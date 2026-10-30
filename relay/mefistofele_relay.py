#!/usr/bin/env python3
"""MEFISTOFELE DeckWhisperer - relay per l'accesso remoto.

Gira sul server (VPS). Il PC del relatore si collega IN USCITA (nessuna porta da aprire sul PC),
i telefoni aprono  https://<server>:<porta>/s/<codice>/  e il relay inoltra al PC solo le richieste
consentite (pagina, stato, anteprime, comandi). Solo libreria standard di Python (3.9 o successivo).

  python3 mefistofele_relay.py --cert fullchain.pem --key privkey.pem --porta 8443

Il PC si registra con un codice casuale (>= 22 caratteri) e una chiave segreta: la chiave e' tenuta
solo come impronta SHA-256, serve a impedire che qualcun altro prenda il posto del PC.
"""
import argparse
import base64
import hashlib
import hmac
import json
import re
import ssl
import sys
import threading
import time
from collections import deque
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlsplit, parse_qs, quote

VERSION = "1.0"
SID_RE = re.compile(r"^[A-Za-z0-9_-]{22,64}$")
RID_RE = re.compile(r"^[0-9]{1,12}$")
# percorsi che un telefono puo' chiedere al PC (tutto il resto e' rifiutato)
GET_OK = (re.compile(r"^/$"), re.compile(r"^/guida$"), re.compile(r"^/favicon\.ico$"), re.compile(r"^/qr\.svg$"),
          re.compile(r"^/static/[A-Za-z0-9_.-]{1,60}$"), re.compile(r"^/slide/[A-Za-z0-9_-]{1,32}/\d{1,5}\.png$"))
MAX_SESSIONS = 300
SESSION_IDLE = 300          # secondi senza contatti dal PC prima di chiudere la sessione
PC_ONLINE = 45              # il PC e' "online" se ha chiesto lavoro negli ultimi N secondi
POLL_WAIT = 25
FORWARD_TIMEOUT = 12
MAX_PC_BODY = 6 * 1024 * 1024
MAX_PHONE_BODY = 4096
MAX_SSE_PER_SESSION = 60
CACHE_SLIDE = 300
CACHE_STATIC = 3600
CACHE_MAX_BYTES = 24 * 1024 * 1024


def log(msg):
    sys.stderr.write("%s %s\n" % (time.strftime("%Y-%m-%d %H:%M:%S"), msg))
    sys.stderr.flush()


class Session:
    def __init__(self, sid, keyhash):
        self.sid = sid
        self.keyhash = keyhash
        self.lock = threading.Condition()
        self.last_pc = time.time()
        self.created = time.time()
        self.state = None            # ultimo stato (stringa JSON)
        self.state_ver = 0
        self.jobs = deque()          # richieste in attesa del PC
        self.waiting = {}            # id -> job
        self.next_id = 1
        self.cache = {}              # chiave -> (scadenza, status, ctype, corpo)
        self.cache_bytes = 0
        self.sse = 0
        self.closed = False
        self.requests = 0

    def online(self):
        return not self.closed and time.time() - self.last_pc < PC_ONLINE


SESSIONS = {}
SLOCK = threading.Lock()
RATE = {}
RATE_LOCK = threading.Lock()


def rate_ok(ip, limit=60, per=10.0):
    """Limite semplice di richieste per indirizzo (finestra scorrevole)."""
    now = time.time()
    with RATE_LOCK:
        q = RATE.setdefault(ip, deque())
        while q and now - q[0] > per:
            q.popleft()
        if len(q) >= limit:
            return False
        q.append(now)
        if len(RATE) > 5000:
            for k in [k for k, v in RATE.items() if not v or now - v[-1] > per]:
                RATE.pop(k, None)
    return True


def get_session(sid):
    with SLOCK:
        return SESSIONS.get(sid)


def drop_session(s):
    with SLOCK:
        if SESSIONS.get(s.sid) is s:
            del SESSIONS[s.sid]
    with s.lock:
        s.closed = True
        for job in list(s.waiting.values()):
            job["event"].set()
        s.lock.notify_all()


def cleaner():
    while True:
        time.sleep(20)
        now = time.time()
        with SLOCK:
            old = [s for s in SESSIONS.values() if now - s.last_pc > SESSION_IDLE]
        for s in old:
            log("sessione scaduta %s..." % s.sid[:6])
            drop_session(s)


def keyhash(key):
    return hashlib.sha256(key.encode("utf-8")).hexdigest()


class Handler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"
    server_version = "MefistofeleRelay/" + VERSION
    timeout = 60                  # connessioni inattive (lettura) chiuse dopo 60 secondi

    def setup(self):
        super().setup()
        if isinstance(self.request, ssl.SSLSocket):
            self.request.do_handshake()      # gia' con il timeout di setup(): un client lento non blocca gli altri

    def log_message(self, *a):
        pass

    def ip(self):
        return self.client_address[0]

    # ---- risposte
    def send_body(self, code, ctype, body, cache="no-store", extra=None):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", cache)
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Referrer-Policy", "no-referrer")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        if self.command != "HEAD":
            self.wfile.write(body)

    def text(self, code, msg):
        self.send_body(code, "text/plain; charset=utf-8", msg.encode("utf-8"))

    def read_body(self, limit):
        try:
            n = int(self.headers.get("Content-Length") or 0)
        except ValueError:
            return None
        if n < 0 or n > limit:
            return None
        return self.rfile.read(n) if n else b""

    # ---- instradamento
    def do_GET(self):
        self.route("GET")

    def do_HEAD(self):
        self.route("GET")

    def do_POST(self):
        self.route("POST")

    def route(self, method):
        parts = urlsplit(self.path)
        path = parts.path
        query = parse_qs(parts.query)
        try:
            if path == "/health":
                with SLOCK:
                    n = len(SESSIONS)
                self.send_body(200, "application/json", json.dumps({"ok": True, "version": VERSION, "sessions": n}).encode())
            elif path == "/":
                self.send_body(200, "text/html; charset=utf-8", LANDING.encode("utf-8"), "public, max-age=300")
            elif path.startswith("/pc/"):
                self.pc(method, path, query)
            elif path.startswith("/s/"):
                self.phone(method, path, query, parts.query)
            else:
                self.text(404, "Not found")
        except (BrokenPipeError, ConnectionResetError, ConnectionAbortedError, ssl.SSLError, OSError):
            pass
        except Exception as e:
            log("errore %s %s: %r" % (method, path[:80], e))
            try:
                self.text(500, "Server error")
            except Exception:
                pass

    # ---- lato PC
    def pc(self, method, path, query):
        seg = path.split("/")          # ['', 'pc', sid, azione, ...]
        if len(seg) < 4 or not SID_RE.match(seg[2]):
            self.text(404, "Not found")
            return
        sid, action = seg[2], seg[3]
        key = self.headers.get("X-DW-Key", "")
        if len(key) < 16 or len(key) > 128:
            self.text(403, "key")
            return
        if not rate_ok("pc:" + self.ip(), 400, 10.0):
            self.text(429, "slow down")
            return
        kh = keyhash(key)
        s = get_session(sid)
        if s is None:
            if action == "bye":
                self.send_body(200, "text/plain", b"ok")
                return
            with SLOCK:
                if len(SESSIONS) >= MAX_SESSIONS:
                    self.text(503, "busy")
                    return
                s = SESSIONS.setdefault(sid, Session(sid, kh))
            log("nuova sessione %s... da %s" % (sid[:6], self.ip()))
        if not hmac.compare_digest(s.keyhash, kh):
            self.text(403, "key")
            return
        s.last_pc = time.time()
        if action == "state" and method == "POST":
            body = self.read_body(64 * 1024)
            if body is None:
                self.text(413, "too big")
                return
            try:
                json.loads(body.decode("utf-8"))
            except Exception:
                self.text(400, "json")
                return
            with s.lock:
                s.state = body.decode("utf-8")
                s.state_ver += 1
                s.lock.notify_all()
            self.send_body(200, "text/plain", b"ok")
        elif action == "poll" and method == "GET":
            job = None
            end = time.time() + POLL_WAIT
            with s.lock:
                while not s.closed:
                    if s.jobs:
                        job = s.jobs.popleft()
                        break
                    left = end - time.time()
                    if left <= 0:
                        break
                    s.lock.wait(timeout=min(left, 5))
                    s.last_pc = time.time()
            if job is None:
                self.send_response(204)
                self.send_header("Cache-Control", "no-store")
                self.end_headers()
                return
            out = {"id": job["id"], "method": job["method"], "path": job["path"]}
            if job["body"]:
                out["body"] = base64.b64encode(job["body"]).decode("ascii")
            self.send_body(200, "application/json", json.dumps(out).encode("utf-8"))
        elif action == "resp" and method == "POST" and len(seg) >= 5 and RID_RE.match(seg[4]):
            body = self.read_body(MAX_PC_BODY)
            if body is None:
                self.text(413, "too big")
                return
            try:
                status = int((query.get("status") or ["502"])[0])
            except ValueError:
                status = 502
            if status < 200 or status > 599:
                status = 502
            ctype = self.headers.get("Content-Type", "application/octet-stream")[:100]
            with s.lock:
                job = s.waiting.get(int(seg[4]))
                if job is not None:
                    job["resp"] = (status, ctype, body)
                    job["event"].set()
            self.send_body(200, "text/plain", b"ok")
        elif action == "bye" and method == "POST":
            drop_session(s)
            self.send_body(200, "text/plain", b"ok")
        else:
            self.text(404, "Not found")

    # ---- lato telefono
    def phone(self, method, path, query, rawquery):
        seg = path.split("/", 3)       # ['', 's', sid, resto]
        if len(seg) < 3 or not SID_RE.match(seg[2]):
            self.text(404, "Not found")
            return
        sid = seg[2]
        rest = "/" + (seg[3] if len(seg) > 3 else "")
        if len(seg) == 3:              # /s/<sid> senza barra: riporta alla pagina
            self.send_response(301)
            self.send_header("Location", "/s/%s/" % sid)
            self.send_header("Content-Length", "0")
            self.end_headers()
            return
        if not rate_ok(self.ip(), 240, 10.0):
            self.text(429, "slow down")
            return
        s = get_session(sid)
        if s is None or not s.online():
            self.send_body(503, "text/html; charset=utf-8", OFFLINE.encode("utf-8"))
            return
        if method == "GET" and rest == "/events":
            self.events(s)
            return
        if method == "GET" and rest == "/state":
            with s.lock:
                data = s.state
            if data is None:
                self.text(503, "no state")
            else:
                self.send_body(200, "application/json; charset=utf-8", data.encode("utf-8"))
            return
        if method == "POST" and rest == "/cmd":
            body = self.read_body(MAX_PHONE_BODY)
            if body is None:
                self.text(413, "too big")
                return
            if not rate_ok("cmd:" + self.ip(), 60, 10.0):
                self.text(429, "slow down")
                return
            self.forward(s, "POST", "/cmd", body, None)
            return
        if method == "GET" and any(r.match(rest) for r in GET_OK):
            ttl = CACHE_SLIDE if rest.startswith("/slide/") else CACHE_STATIC if rest.startswith("/static/") or rest == "/favicon.ico" else 0
            full = rest
            if rest == "/guida" and "lang" in query:
                lang = (query["lang"][0] or "")[:2]
                if re.match(r"^[a-z]{2}$", lang):
                    full = "/guida?lang=" + lang
            if rest.startswith("/slide/") and query.get("v"):
                full = rest + "?v=" + quote((query["v"][0] or "")[:100], safe="")
            self.forward(s, "GET", full, b"", ttl)
            return
        self.text(404, "Not found")

    def forward(self, s, method, path, body, ttl):
        ck = method + path
        if ttl:
            with s.lock:
                hit = s.cache.get(ck)
            if hit and hit[0] > time.time():
                self.send_body(hit[1], hit[2], hit[3], "private, max-age=%d" % min(ttl, 300))
                return
        job = {"event": threading.Event(), "resp": None, "method": method, "path": path, "body": body}
        with s.lock:
            if s.closed:
                self.text(503, "offline")
                return
            job["id"] = s.next_id
            s.next_id += 1
            s.waiting[job["id"]] = job
            s.jobs.append(job)
            s.requests += 1
            s.lock.notify_all()
        try:
            job["event"].wait(FORWARD_TIMEOUT)
        finally:
            with s.lock:
                s.waiting.pop(job["id"], None)
                try:
                    s.jobs.remove(job)
                except ValueError:
                    pass
        if job["resp"] is None:
            self.text(504, "The presenter's PC did not answer")
            return
        status, ctype, data = job["resp"]
        if ttl and status == 200:
            with s.lock:
                if s.cache_bytes + len(data) < CACHE_MAX_BYTES:
                    s.cache[ck] = (time.time() + ttl, status, ctype, data)
                    s.cache_bytes += len(data)
        self.send_body(status, ctype, data, ("private, max-age=%d" % min(ttl, 300)) if ttl and status == 200 else "no-store")

    def events(self, s):
        with s.lock:
            if s.sse >= MAX_SSE_PER_SESSION:
                self.text(503, "too many viewers")
                return
            s.sse += 1
        try:
            self.send_response(200)
            self.send_header("Content-Type", "text/event-stream; charset=utf-8")
            self.send_header("Cache-Control", "no-store")
            self.send_header("X-Accel-Buffering", "no")
            self.send_header("Connection", "keep-alive")
            self.end_headers()
            last = -1
            while True:
                with s.lock:
                    if s.closed or not s.online():
                        break
                    if s.state_ver == last:
                        s.lock.wait(timeout=15)
                    ver, data = s.state_ver, s.state
                    gone = s.closed or not s.online()
                if gone:
                    break
                if ver != last and data is not None:
                    last = ver
                    self.wfile.write(("data: " + data + "\n\n").encode("utf-8"))
                else:
                    self.wfile.write(b": ping\n\n")
                self.wfile.flush()
        except (BrokenPipeError, ConnectionResetError, ConnectionAbortedError, ssl.SSLError, OSError):
            pass
        finally:
            with s.lock:
                s.sse -= 1
            self.close_connection = True


PAGE_CSS = "body{font-family:system-ui,Segoe UI,Arial,sans-serif;background:#f4f6f7;color:#14262b;margin:0;display:grid;place-items:center;min-height:100vh}main{max-width:460px;padding:24px;text-align:center}h1{font-size:22px}p{color:#5b6b70;line-height:1.5}"
LANDING = ("<!DOCTYPE html><html lang=en><meta charset=utf-8><meta name=viewport content='width=device-width,initial-scale=1'>"
           "<title>MEFISTOFELE DeckWhisperer</title><meta name=robots content=noindex><style>%s</style><main><h1>MEFISTOFELE DeckWhisperer</h1>"
           "<p>Remote access relay. Open the link or scan the QR code shown by the presenter.</p></main></html>") % PAGE_CSS
OFFLINE = ("<!DOCTYPE html><html lang=en><meta charset=utf-8><meta name=viewport content='width=device-width,initial-scale=1'>"
           "<title>MEFISTOFELE DeckWhisperer</title><meta name=robots content=noindex><style>%s</style><main><h1>Presenter offline</h1>"
           "<p>The presenter's PC is not connected (or remote access was switched off). Check the link and try again in a moment.</p></main></html>") % PAGE_CSS


class Server(ThreadingHTTPServer):
    daemon_threads = True
    request_queue_size = 128
    ctx = None                    # contesto TLS (None = HTTP semplice, solo prove)

    def handle_error(self, request, client_address):
        exc = sys.exc_info()[1]
        if not isinstance(exc, (OSError, ssl.SSLError)):     # i client che cadono durante l'handshake sono normali
            log("errore connessione da %s: %r" % (client_address[0], exc))

    def get_request(self):
        """Accetta la connessione senza bloccarsi: l'handshake TLS lo fa il thread della richiesta."""
        sock, addr = self.socket.accept()
        if self.ctx is not None:
            sock = self.ctx.wrap_socket(sock, server_side=True, do_handshake_on_connect=False)
        return sock, addr


def main():
    ap = argparse.ArgumentParser(description="Relay MEFISTOFELE DeckWhisperer")
    ap.add_argument("--porta", type=int, default=8443)
    ap.add_argument("--host", default="0.0.0.0")
    ap.add_argument("--cert", help="certificato HTTPS (fullchain.pem)")
    ap.add_argument("--key", help="chiave privata HTTPS (privkey.pem)")
    ap.add_argument("--http", action="store_true", help="senza HTTPS (solo prove locali)")
    a = ap.parse_args()
    srv = Server((a.host, a.porta), Handler)
    if not a.http:
        if not (a.cert and a.key):
            ap.error("servono --cert e --key (oppure --http per le prove)")
        ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
        ctx.minimum_version = ssl.TLSVersion.TLSv1_2
        ctx.load_cert_chain(a.cert, a.key)
        srv.ctx = ctx
    threading.Thread(target=cleaner, daemon=True).start()
    log("relay %s in ascolto su %s:%d (%s)" % (VERSION, a.host, a.porta, "http" if a.http else "https"))
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
