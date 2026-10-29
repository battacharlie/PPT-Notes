#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
reMNote - note del relatore da PowerPoint
Legge le note del relatore da PowerPoint (in presentazione) e le pubblica su una
pagina web raggiungibile dalla rete locale, in tempo reale.

Funzioni:
  - piu' presentazioni aperte contemporaneamente: ciascun dispositivo sceglie quale seguire
  - comandi per avanzare/tornare/andare a una slide dal telefono (PIN opzionale)
  - schermo sempre acceso (dove il browser lo consente)

Uso:
    py remnote.py                  (Windows, con PowerPoint installato)
    py remnote.py --pin 1234       (i comandi richiedono il PIN)
    py remnote.py --no-control     (disabilita del tutto i comandi)
    py remnote.py --demo           (prova senza PowerPoint)
    py remnote.py --port 8080
"""
import argparse
import json
import os
import queue
import re
import socket
import sys
import tempfile
import threading
import time
import webbrowser
from collections import OrderedDict, deque
from html import escape as html_escape
from xml.sax.saxutils import escape
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler

APP_NAME = "reMNote"
try:
    from _buildinfo import APP_VERSION  # scritto dalla compilazione (tools/buildinfo.py)
except Exception:
    try:  # avvio da sorgente: legge il file VERSION della repo
        with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "VERSION"), encoding="utf-8") as _f:
            APP_VERSION = _f.read().strip()
    except Exception:
        APP_VERSION = "sviluppo"
PIN = ""
CONTROL = True
BASE_URL = ""

ASSET_FILES = {
    "remnote-512.png": "image/png", "remnote-192.png": "image/png", "remnote-64.png": "image/png",
    "remnote-32.png": "image/png", "apple-touch-icon.png": "image/png", "mwtl-logo.png": "image/png",
    "remnote.ico": "image/x-icon",
}


def resource_path(*parts):
    """Percorso di un file incluso nel programma (anche dentro l'exe PyInstaller)."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, *parts)


CREDITS_HTML = (
    '<b>reMNote</b> – Note del relatore<br>Ideato e sviluppato da '
    '<a href="https://www.andreabattaglioli.it" target="_blank" rel="noopener">www.andreabattaglioli.it</a> · '
    '<a class="mwtl" href="https://www.mywebtvlive.com" target="_blank" rel="noopener">'
    '<img src="/static/mwtl-logo.png" alt="" width="18" height="18">MyWebTvLive.com</a> · Made in Bolognina'
)

STATE = {"pin_required": False, "control": True, "shows": []}
VERSION = 0
COND = threading.Condition()
CMDQ = queue.Queue()
IMGQ = queue.Queue()          # richieste di anteprima slide (eseguite dal thread che usa COM)
SLIDE_CACHE = OrderedDict()   # (id presentazione, numero slide) -> immagine PNG
SLIDE_LOCK = threading.Lock()
SLIDE_CACHE_MAX = 40
SLIDE_RE = re.compile(r"^/slide/([A-Za-z0-9_-]{1,32})/(\d{1,5})\.png$")
DEMO_MODE = False


def set_state(shows):
    """Aggiorna lo stato e avvisa i client solo se qualcosa e' cambiato."""
    global VERSION
    new = {"pin_required": bool(PIN), "control": CONTROL, "shows": shows}
    with COND:
        if new != STATE:
            STATE.clear()
            STATE.update(new)
            VERSION += 1
            COND.notify_all()


def snapshot():
    with COND:
        return VERSION, json.dumps(STATE, ensure_ascii=False)


# ---------------------------------------------------------------- PowerPoint
def notes_of(slide):
    """Testo del segnaposto 'note' (corpo) della pagina note di una slide."""
    try:
        for shp in slide.NotesPage.Shapes:
            # 14 = msoPlaceholder, 2 = ppPlaceholderBody
            if shp.Type == 14 and shp.PlaceholderFormat.Type == 2:
                if shp.HasTextFrame:
                    return (shp.TextFrame.TextRange.Text
                            .replace("\r\n", "\n").replace("\r", "\n").replace("\x0b", "\n"))
    except Exception:
        pass
    return ""


def win_id(win, i):
    try:
        return str(win.HWND)
    except Exception:
        return "w%d" % i


def read_shows(app):
    shows = []
    for i in range(1, int(app.SlideShowWindows.Count) + 1):
        try:
            win = app.SlideShowWindows(i)
            view = win.View
            idx = int(view.CurrentShowPosition)
            pres = win.Presentation
            total = int(pres.Slides.Count)
            nxt = notes_of(pres.Slides(idx + 1)) if idx < total else ""
            shows.append({
                "id": win_id(win, i),
                "title": str(pres.Name),
                "index": idx,
                "total": total,
                "notes": notes_of(view.Slide),
                "next_notes": nxt,
            })
        except Exception:
            continue
    return shows


def run_command(app, cmd):
    wid = cmd.get("id")
    act = cmd.get("action")
    for i in range(1, int(app.SlideShowWindows.Count) + 1):
        win = app.SlideShowWindows(i)
        if wid and win_id(win, i) != wid:
            continue
        view = win.View
        if act == "next":
            view.Next()
        elif act == "prev":
            view.Previous()
        elif act == "first":
            view.First()
        elif act == "last":
            view.Last()
        elif act == "goto":
            n = int(cmd.get("n"))
            total = int(win.Presentation.Slides.Count)
            if 1 <= n <= total:
                view.GotoSlide(n)
        return True
    return False


LOG_SEEN = set()
LOG_LINES = deque(maxlen=60)   # ultimi messaggi, mostrati anche nella pagina /diagnostica
LOG_FILE_USED = [""]


def log_path():
    base = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "reMNote.log")


def log_paths():
    """Dove prova a scrivere il registro: accanto all'exe, poi nella cartella temporanea."""
    return [log_path(), os.path.join(tempfile.gettempdir(), "reMNote.log")]


def log(msg):
    """Scrive un messaggio in console, in memoria e nel file reMNote.log (ogni messaggio una volta sola)."""
    if msg in LOG_SEEN or len(LOG_SEEN) > 50:
        return
    LOG_SEEN.add(msg)
    line = "%s %s" % (time.strftime("%Y-%m-%d %H:%M:%S"), msg)
    LOG_LINES.append(line)
    try:
        print(" [registro] " + msg)
    except Exception:
        pass
    for pth in log_paths():
        try:
            with open(pth, "a", encoding="utf-8") as f:
                f.write(line + "\n")
            LOG_FILE_USED[0] = pth
            break
        except Exception:
            continue


def diagnostica_info():
    """Testo con le informazioni utili per capire un problema, senza cercare file."""
    lines = [
        "reMNote %s" % APP_VERSION,
        "Programma: %s%s" % (sys.executable, "  (exe)" if getattr(sys, "frozen", False) else "  (da sorgente)"),
        "Indirizzo: %s" % BASE_URL,
        "Registro previsto: %s" % log_path(),
        "Registro scritto in: %s" % (LOG_FILE_USED[0] or "(non ancora scritto o scrittura impossibile)"),
        "Modalita demo: %s" % ("si" if DEMO_MODE else "no"),
        "",
        "Ultimi messaggi del registro:",
    ] + (list(LOG_LINES) or ["(nessuno)"])
    return "\n".join(lines)


def diagnostica_page():
    """Pagina /diagnostica."""
    return DIAG_PAGE.replace("__INFO__", html_escape(diagnostica_info()))


DIAG_PAGE = r"""<!DOCTYPE html><html lang="it"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1"><title>reMNote - diagnostica</title>
<link rel="icon" type="image/png" href="/static/remnote-64.png">
<style>body{font:14px Arial,sans-serif;margin:16px;background:#fff;color:#111}
pre{background:#f3f3f3;padding:10px;border-radius:6px;white-space:pre-wrap;overflow-wrap:anywhere}
button{font:15px Arial,sans-serif;padding:8px 14px;margin:4px 6px 4px 0;cursor:pointer}</style></head>
<body><h2>reMNote - diagnostica</h2>
<pre id="info">__INFO__</pre>
<p><button id="prova">Prova le anteprime ora</button><button id="copia">Copia tutto</button></p>
<pre id="out">Premi "Prova le anteprime ora" con una presentazione in corso.</pre>
<script>
async function prova() {
  const out = document.getElementById("out"); out.textContent = "Prova in corso…";
  const L = [];
  try {
    const st = await (await fetch("/state")).json();
    L.push("Presentazioni in corso: " + st.shows.length);
    for (const s of st.shows) {
      L.push("- id=" + s.id + "  titolo=" + s.title + "  slide=" + s.index + "/" + s.total);
      for (const n of [s.index, s.index + 1]) {
        if (n > s.total) continue;
        const url = "/slide/" + encodeURIComponent(s.id) + "/" + n + ".png";
        const t0 = performance.now();
        const r = await fetch(url);
        const ct = r.headers.get("Content-Type") || "";
        const body = ct.indexOf("image") === 0 ? (await r.blob()).size + " byte" : await r.text();
        L.push("    " + url + "  ->  " + r.status + "  " + ct + "  " + body + "  (" + Math.round(performance.now() - t0) + " ms)");
      }
    }
    if (!st.shows.length) L.push("Nessuna presentazione in corso: avvia una presentazione in PowerPoint e riprova.");
  } catch (e) { L.push("Errore: " + e); }
  out.textContent = L.join("\n");
  fetch("/diagnostica?solo=info").then(r => r.text()).then(t => { document.getElementById("info").textContent = t; }).catch(() => {});
}
document.getElementById("prova").onclick = prova;
document.getElementById("copia").onclick = async () => {
  const txt = document.getElementById("info").textContent + "\n\n" + document.getElementById("out").textContent;
  try { await navigator.clipboard.writeText(txt); alert("Copiato: incollalo nella chat."); }
  catch (e) {
    const r = document.createRange(); r.selectNodeContents(document.body);
    const sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(r);
    alert("Selezionato tutto: premi Ctrl+C e incolla nella chat.");
  }
};
</script></body></html>"""


def cache_get(key):
    with SLIDE_LOCK:
        v = SLIDE_CACHE.get(key)
        if v is not None:
            SLIDE_CACHE.move_to_end(key)
        return v


def cache_put(key, data):
    with SLIDE_LOCK:
        SLIDE_CACHE[key] = data
        SLIDE_CACHE.move_to_end(key)
        while len(SLIDE_CACHE) > SLIDE_CACHE_MAX:
            SLIDE_CACHE.popitem(last=False)


def cache_clear():
    with SLIDE_LOCK:
        SLIDE_CACHE.clear()


def export_slide(app, job):
    """Esporta una slide in PNG (va chiamata solo dal thread che usa COM).
    Se fallisce, il motivo resta in job["error"] e viene mostrato sulla pagina."""
    key = (job["id"], job["index"])
    data = cache_get(key)
    job["error"] = ""
    if data is None:
        job["error"] = "presentazione non trovata (id %s)" % job["id"]
        for i in range(1, int(app.SlideShowWindows.Count) + 1):
            win = app.SlideShowWindows(i)
            if win_id(win, i) != job["id"]:
                continue
            pres = win.Presentation
            total = int(pres.Slides.Count)
            if not 1 <= job["index"] <= total:
                job["error"] = "slide %d fuori intervallo (1-%d)" % (job["index"], total)
                break
            w = 960
            try:
                h = int(round(w * float(pres.PageSetup.SlideHeight) / float(pres.PageSetup.SlideWidth)))
            except Exception:
                h = 540
            path = os.path.join(tempfile.gettempdir(), "reMNote_%d_%d.png" % (os.getpid(), job["index"]))
            errs = []
            try:
                # tentativo 1: dalla presentazione; tentativo 2: dalla slide in onda della finestra di presentazione
                for attempt in (1, 2):
                    try:
                        if attempt == 1:
                            sl = pres.Slides(job["index"])
                        else:
                            if int(win.View.CurrentShowPosition) != job["index"]:
                                break
                            sl = win.View.Slide
                        sl.Export(path, "PNG", w, h)
                        if os.path.exists(path) and os.path.getsize(path) > 0:
                            with open(path, "rb") as f:
                                data = f.read()
                            break
                        errs.append("file immagine non creato")
                    except Exception as e:
                        errs.append("%s: %s" % (type(e).__name__, e))
            finally:
                try:
                    os.remove(path)
                except OSError:
                    pass
            job["error"] = "" if data else "esportazione fallita (%s)" % (" | ".join(errs) or "nessun tentativo possibile")
            break
        if data:
            cache_put(key, data)
    job["result"] = data
    if not data:
        log("Anteprima slide %s/%s: %s" % (job["id"], job["index"], job["error"]))
    job["event"].set()


def process_image_jobs(app):
    while True:
        try:
            job = IMGQ.get_nowait()
        except queue.Empty:
            break
        if time.time() - job["t"] > 8.0:
            job["event"].set()   # nessuno la sta piu' aspettando
            continue
        try:
            export_slide(app, job)
        except Exception as e:
            job["result"] = None
            job["error"] = "%s: %s" % (type(e).__name__, e)
            log("Anteprima slide: " + job["error"])
            job["event"].set()


def fail_image_jobs():
    while True:
        try:
            job = IMGQ.get_nowait()
        except queue.Empty:
            break
        job["event"].set()


def poll_powerpoint(interval):
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    app = None
    while True:
        shows = []
        try:
            if app is None:
                app = win32com.client.GetActiveObject("PowerPoint.Application")
            while True:
                try:
                    cmd = CMDQ.get_nowait()
                except queue.Empty:
                    break
                if time.time() - cmd.get("t", 0) > 2.0:
                    continue  # comando vecchio: non va eseguito in ritardo
                try:
                    run_command(app, cmd)
                except Exception:
                    pass
            process_image_jobs(app)
            if int(app.SlideShowWindows.Count) > 0:
                shows = read_shows(app)
            else:
                cache_clear()
        except Exception:
            # PowerPoint chiuso o COM momentaneamente occupato: ricollega al giro dopo
            app = None
            fail_image_jobs()
            while not CMDQ.empty():
                try:
                    CMDQ.get_nowait()
                except queue.Empty:
                    break
        set_state(shows)
        time.sleep(interval)


DEMO_SHOWS = [
    {"id": "d1", "title": "demo_apertura.pptx", "index": 1, "notes": [
        "Benvenuti.\nApro l'evento presentando gli ospiti.",
        "Prima slide dati: sottolineare il +12% rispetto all'anno scorso.",
        "Pausa di 10 secondi, poi passare al video.",
        "Ringraziamenti finali e domande dal pubblico."]},
    {"id": "d2", "title": "demo_relatore2.pptx", "index": 1, "notes": [
        "Presentazione del secondo relatore.",
        "Case history: tre punti chiave.",
        "Conclusioni."]},
]


def demo_slide(sid, index):
    """Anteprima finta (SVG) per la modalita' demo."""
    for d in DEMO_SHOWS:
        if d["id"] == sid and 1 <= index <= len(d["notes"]):
            line = d["notes"][index - 1].split("\n")[0][:70]
            svg = (
                '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 960 540">'
                '<rect width="960" height="540" fill="#14434d"/>'
                '<text x="480" y="150" fill="#f26a21" font-family="Arial" font-size="36" text-anchor="middle">%s</text>'
                '<text x="480" y="300" fill="#ffffff" font-family="Arial" font-size="96" font-weight="700" '
                'text-anchor="middle">Slide %d / %d</text>'
                '<text x="480" y="420" fill="#cfe" font-family="Arial" font-size="30" text-anchor="middle">%s</text>'
                '</svg>' % (escape(d["title"]), index, len(d["notes"]), escape(line)))
            return svg.encode("utf-8"), "image/svg+xml"
    return None, ""


def poll_demo(_interval):
    while True:
        while True:
            try:
                cmd = CMDQ.get_nowait()
            except queue.Empty:
                break
            for d in DEMO_SHOWS:
                if cmd.get("id") and d["id"] != cmd["id"]:
                    continue
                n = len(d["notes"])
                act = cmd.get("action")
                if act == "next":
                    d["index"] = min(n, d["index"] + 1)
                elif act == "prev":
                    d["index"] = max(1, d["index"] - 1)
                elif act == "first":
                    d["index"] = 1
                elif act == "last":
                    d["index"] = n
                elif act == "goto" and 1 <= int(cmd.get("n")) <= n:
                    d["index"] = int(cmd["n"])
                break
        shows = []
        for d in DEMO_SHOWS:
            n = len(d["notes"])
            i = d["index"]
            shows.append({"id": d["id"], "title": d["title"], "index": i, "total": n,
                          "notes": d["notes"][i - 1],
                          "next_notes": d["notes"][i] if i < n else ""})
        set_state(shows)
        time.sleep(0.1)


# ------------------------------------------------------------------------ QR
def qr_matrix(text):
    """Matrice del QR (True = modulo scuro), senza bordo."""
    import qrcode
    qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_M, border=0)
    qr.add_data(text)
    qr.make(fit=True)
    return qr.get_matrix()


def qr_svg(text):
    m = qr_matrix(text)
    n = len(m)
    b = 4  # zona di quiete
    d = []
    for y, row in enumerate(m):
        for x, v in enumerate(row):
            if v:
                d.append("M%d %dh1v1h-1z" % (x + b, y + b))
    size = n + 2 * b
    return (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 %d %d" shape-rendering="crispEdges">'
        '<rect width="100%%" height="100%%" fill="#fff"/><path d="%s" fill="#000"/></svg>'
        % (size, size, "".join(d))
    )


def qr_console(text):
    """QR a blocchi per la console: i moduli chiari sono i caratteri pieni."""
    m = qr_matrix(text)
    b = 2
    n = len(m)
    size = n + 2 * b

    def dark(x, y):
        xx, yy = x - b, y - b
        return 0 <= xx < n and 0 <= yy < n and m[yy][xx]

    lines = []
    for y in range(0, size, 2):
        line = ""
        for x in range(size):
            top_lit = not dark(x, y)
            bot_lit = not dark(x, y + 1) if y + 1 < size else True
            line += "█" if top_lit and bot_lit else "▀" if top_lit else "▄" if bot_lit else " "
        lines.append(line)
    return "\n".join(lines)


QR_PAGE = """<!DOCTYPE html><html lang="it"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1"><title>QR note relatore</title>
<link rel="icon" type="image/png" href="/static/remnote-64.png">
<style>body{margin:0;min-height:100vh;display:flex;flex-direction:column;align-items:center;
justify-content:center;background:#fff;font-family:Arial,sans-serif;color:#000;text-align:center}
img.qr{width:min(70vw,70vh);height:min(70vw,70vh)}p{font-size:28px;margin:10px}
.hd{display:flex;align-items:center;justify-content:center}
.hd img{width:130px;height:130px}.cr{font-size:14px;color:#444;line-height:1.5}
.cr a{color:#14434d}.cr img{vertical-align:middle;margin-right:4px}</style></head>
<body><div class="hd"><img src="/static/remnote-192.png" alt="reMNote"></div>
<img class="qr" src="/qr.svg" alt="QR"><p>__URL__</p>
<p style="font-size:16px">Inquadra con lo smartphone (stessa rete Wi-Fi)</p>
<div class="cr">__CREDITS__</div></body></html>"""


# -------------------------------------------------------------------- Server
class Handler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"

    def log_message(self, *a):
        pass

    def _send(self, code, ctype, body):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _json(self, code, obj):
        self._send(code, "application/json; charset=utf-8", json.dumps(obj).encode("utf-8"))

    def _slide(self, sid, index):
        ctype = "image/png"
        job = {"error": ""}
        if DEMO_MODE:
            data, ctype = demo_slide(sid, index)
        else:
            data = cache_get((sid, index))
            if data is None:
                job = {"id": sid, "index": index, "t": time.time(),
                       "event": threading.Event(), "result": None, "error": ""}
                IMGQ.put(job)
                job["event"].wait(6)
                data = job["result"]
                if not data and not job["error"]:
                    job["error"] = "PowerPoint non ha risposto entro 6 secondi"
                    log("Anteprima slide %s/%s: %s" % (sid, index, job["error"]))
        if not data:
            msg = (job["error"] if not DEMO_MODE else "") or "Anteprima non disponibile"
            self._send(404, "text/plain; charset=utf-8", msg[:300].encode("utf-8"))
            return
        self.send_response(200)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "private, max-age=300")
        self.end_headers()
        self.wfile.write(data)

    def do_POST(self):
        path = self.path.split("?")[0]
        try:
            length = int(self.headers.get("Content-Length") or 0)
        except ValueError:
            length = 0
        raw = self.rfile.read(min(length, 4096)) if length > 0 else b""
        if path != "/cmd":
            self._json(404, {"ok": False, "error": "non trovato"})
            return
        if not CONTROL:
            self._json(403, {"ok": False, "error": "comandi disabilitati"})
            return
        try:
            body = json.loads(raw.decode("utf-8") or "{}")
        except Exception:
            self._json(400, {"ok": False, "error": "richiesta non valida"})
            return
        if PIN and str(body.get("pin", "")) != PIN:
            self._json(403, {"ok": False, "error": "pin"})
            return
        action = body.get("action")
        if action not in ("next", "prev", "first", "last", "goto"):
            self._json(400, {"ok": False, "error": "azione sconosciuta"})
            return
        cmd = {"id": str(body.get("id", "")), "action": action, "t": time.time()}
        if action == "goto":
            try:
                cmd["n"] = int(body.get("n"))
            except Exception:
                self._json(400, {"ok": False, "error": "numero slide non valido"})
                return
        CMDQ.put(cmd)
        self._json(200, {"ok": True})

    def do_GET(self):
        path = self.path.split("?")[0]
        if path in ("/", "/index.html"):
            self._send(200, "text/html; charset=utf-8",
                       PAGE.replace("__CREDITS__", CREDITS_HTML).replace("__VERSION__", APP_VERSION).encode("utf-8"))
        elif path.startswith("/slide/"):
            m = SLIDE_RE.match(path)
            if not m:
                log("Indirizzo anteprima non riconosciuto (404): %s" % self.path[:150])
                self._send(404, "text/plain; charset=utf-8", b"Non trovato")
                return
            self._slide(m.group(1), int(m.group(2)))
        elif path.startswith("/static/") or path == "/favicon.ico":
            name = "remnote.ico" if path == "/favicon.ico" else path[len("/static/"):]
            if name not in ASSET_FILES:
                self._send(404, "text/plain; charset=utf-8", b"Non trovato")
                return
            try:
                with open(resource_path("assets", name), "rb") as f:
                    data = f.read()
            except OSError:
                self._send(404, "text/plain; charset=utf-8", b"Non trovato")
                return
            self._send(200, ASSET_FILES[name], data)
        elif path == "/diagnostica":
            if "solo=info" in self.path:
                self._send(200, "text/plain; charset=utf-8", diagnostica_info().encode("utf-8"))
            else:
                self._send(200, "text/html; charset=utf-8", diagnostica_page().encode("utf-8"))
        elif path == "/qr.svg":
            self._send(200, "image/svg+xml", qr_svg(BASE_URL).encode("utf-8"))
        elif path == "/qr":
            self._send(200, "text/html; charset=utf-8", QR_PAGE.replace("__URL__", BASE_URL).replace("__CREDITS__", CREDITS_HTML).encode("utf-8"))
        elif path == "/state":
            _, data = snapshot()
            self._send(200, "application/json; charset=utf-8", data.encode("utf-8"))
        elif path == "/events":
            self.send_response(200)
            self.send_header("Content-Type", "text/event-stream; charset=utf-8")
            self.send_header("Cache-Control", "no-store")
            self.send_header("Connection", "keep-alive")
            self.end_headers()
            last = -1
            try:
                while True:
                    with COND:
                        if VERSION == last:
                            COND.wait(timeout=15)
                    ver, data = snapshot()
                    if ver != last:
                        last = ver
                        self.wfile.write(("data: " + data + "\n\n").encode("utf-8"))
                    else:
                        self.wfile.write(b": ping\n\n")
                    self.wfile.flush()
            except (BrokenPipeError, ConnectionResetError, ConnectionAbortedError, OSError):
                pass
        else:
            log("Indirizzo non riconosciuto (404): %s" % self.path[:150])
            self._send(404, "text/plain; charset=utf-8", b"Non trovato")


def local_ips():
    ips = set()
    try:
        for info in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
            ips.add(info[4][0])
    except Exception:
        pass
    return sorted(i for i in ips if not i.startswith("127."))


def primary_ip():
    """IP della scheda usata per uscire in rete (quella giusta per gli smartphone)."""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("10.255.255.255", 1))
        ip = s.getsockname()[0]
        s.close()
        if not ip.startswith("127."):
            return ip
    except Exception:
        pass
    ips = local_ips()
    return ips[0] if ips else "localhost"


PAGE = r"""<!DOCTYPE html>
<html lang="it">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>reMNote</title>
<link rel="icon" type="image/png" sizes="64x64" href="/static/remnote-64.png">
<link rel="icon" type="image/png" sizes="32x32" href="/static/remnote-32.png">
<link rel="apple-touch-icon" href="/static/apple-touch-icon.png">
<meta name="theme-color" content="#14434d">
<style>
  :root { --fg:#ffffff; --bg:#000000; }
  * { box-sizing: border-box; }
  html, body { margin:0; height:100%; background:var(--bg); color:var(--fg); }
  body { display:flex; flex-direction:column; font-family: Arial, sans-serif; }
  #bar { display:flex; flex-wrap:wrap; gap:8px 14px; align-items:center; padding:6px 10px;
         background:#222; color:#eee; font:13px Arial, sans-serif; border-bottom:1px solid #444; }
  #bar label { display:flex; align-items:center; gap:5px; white-space:nowrap; }
  #bar select, #bar input[type=number], #bar button { font:13px Arial, sans-serif; }
  #bar input[type=number] { width:60px; }
  #bar input[type=range] { width:90px; }
  #bar input[type=color] { width:34px; height:24px; padding:0; border:1px solid #777; background:none; }
  #bar button { padding:3px 9px; cursor:pointer; }
  #info { margin-left:auto; display:flex; gap:12px; align-items:center; }
  #dot { width:10px; height:10px; border-radius:50%; background:#c33; display:inline-block; }
  #dot.on { background:#3c3; }
  #main { flex:1; overflow:auto; padding:2vh 3vw; }
  #notes { white-space:pre-wrap; overflow-wrap:anywhere;
           font-family: Arial, sans-serif; font-size:48px; line-height:1.3; letter-spacing:0; }
  #notes.wait { opacity:.5; font-style:italic; }
  #ctl { display:flex; gap:10px; align-items:stretch; padding:8px 10px; background:#222;
         border-top:1px solid #444; color:#eee; font:15px Arial, sans-serif; }
  #ctl button { font:600 20px Arial, sans-serif; padding:12px 8px; cursor:pointer; border-radius:8px;
                border:1px solid #666; background:#333; color:#fff; }
  #ctl .big { flex:1; }
  #ctl .mid { display:flex; flex-direction:column; align-items:center; justify-content:center; min-width:120px; gap:4px; }
  #ctl .mid input { width:70px; font:16px Arial, sans-serif; text-align:center; }
  #ctl .mid button { font-size:14px; padding:4px 10px; }
  #ctlmsg { color:#f88; font-size:12px; min-height:14px; }
  #infoov { display:none; position:fixed; inset:0; background:rgba(0,0,0,.85); z-index:9;
            align-items:center; justify-content:center; padding:16px; }
  #infobox { background:#fff; color:#111; border-radius:14px; padding:22px 26px; max-width:440px;
             text-align:center; font:15px/1.6 Arial, sans-serif; }
  #infobox img.logo { width:150px; height:150px; }
  #infobox .ver { color:#666; font-size:13px; margin-bottom:12px; }
  #infobox a { color:#14434d; }
  #infobox .mwtl img { vertical-align:middle; margin-right:4px; }
  .grp { display:contents; }
  body.nostyle .grp { display:none; }
  #bar button.tg.on { background:#2b8392; color:#fff; border-color:#8fd3df; }
  #slides { display:none; gap:10px; justify-content:center; padding:8px 10px; background:#111;
            border-bottom:1px solid #444; }
  body.showcur #slides, body.shownext #slides { display:flex; }
  #slides figure { display:none; margin:0; flex:1 1 0; max-width:460px; text-align:center; }
  body.showcur #figcur, body.shownext #fignext { display:block; }
  #slides figcaption { font:12px Arial, sans-serif; color:#aaa; margin-bottom:3px; }
  #slides img { display:block; width:100%; max-height:30vh; object-fit:contain; background:#000;
                border:1px solid #555; }
  #slides .ph { display:none; aspect-ratio:16/9; max-height:30vh; width:100%; align-items:center;
                justify-content:center; border:1px dashed #555; color:#999; font:13px Arial, sans-serif;
                padding:6px; text-align:center; overflow:auto; overflow-wrap:anywhere; }
  body.hidebar #bar { display:none; }
  body.noctl #ctl { display:none; }
  #showbar { position:fixed; right:8px; top:8px; display:none; opacity:.4; }
  body.hidebar #showbar { display:block; }
</style>
</head>
<body>
<div id="bar">
  <label id="showlbl">Presentazione
    <select id="show"></select>
  </label>
  <span class="grp">
  <label>Font
    <select id="font">
      <option value="Arial, sans-serif">Arial</option>
      <option value="Verdana, sans-serif">Verdana</option>
      <option value="Tahoma, sans-serif">Tahoma</option>
      <option value="'Trebuchet MS', sans-serif">Trebuchet</option>
      <option value="Calibri, sans-serif">Calibri</option>
      <option value="'Segoe UI', sans-serif">Segoe UI</option>
      <option value="Georgia, serif">Georgia</option>
      <option value="'Times New Roman', serif">Times New Roman</option>
      <option value="'Courier New', monospace">Courier New</option>
    </select>
  </label>
  <label>Dimensione
    <button id="sizeDown" title="Riduci">A-</button>
    <input type="number" id="size" min="10" max="300" step="2">
    <button id="sizeUp" title="Aumenta">A+</button>
  </label>
  <label>Testo <input type="color" id="fg"></label>
  <label>Sfondo <input type="color" id="bg"></label>
  <label>Interlinea <input type="range" id="lh" min="0.9" max="3" step="0.05"> <span id="lhv"></span></label>
  <label>Spaziatura <input type="range" id="ls" min="-2" max="20" step="0.5"> <span id="lsv"></span></label>
  <button id="reset">Ripristina</button>
  </span>
  <button id="full">Schermo intero</button>
  <button id="styleb" class="tg" title="Mostra/nascondi i comandi del testo">Stile testo</button>
  <button id="ctlb" class="tg" title="Mostra/nascondi i pulsanti per avanzare le slide">Comandi</button>
  <button id="curb" class="tg" title="Mostra/nascondi la slide in onda">Slide in onda</button>
  <button id="nextb" class="tg" title="Mostra/nascondi la slide successiva">Slide successiva</button>
  <button id="qrb">QR</button>
  <button id="infob">Info</button>
  <button id="hide">Nascondi barra</button>
  <div id="info"><span id="pos"></span><span id="dot" title="Presentazione"></span></div>
</div>
<div id="slides">
  <figure id="figcur"><figcaption>In onda</figcaption><img id="imgcur" alt="Slide in onda"><div class="ph" id="phcur"></div></figure>
  <figure id="fignext"><figcaption>Successiva</figcaption><img id="imgnext" alt="Slide successiva"><div class="ph" id="phnext"></div></figure>
</div>
<div id="main"><div id="notes" class="wait">In attesa della presentazione…</div></div>
<div id="ctl">
  <button id="prev" class="big">◀ Indietro</button>
  <div class="mid">
    <span id="ctlpos">–</span>
    <span><input type="number" id="gonum" min="1" inputmode="numeric"> <button id="go">Vai</button></span>
    <span id="ctlmsg"></span>
  </div>
  <button id="next" class="big">Avanti ▶</button>
</div>
<button id="showbar">Barra</button>
<div id="infoov">
  <div id="infobox">
    <img class="logo" src="/static/remnote-192.png" alt="reMNote">
    <div class="ver">Versione __VERSION__</div>
    <div>__CREDITS__</div>
    <div style="margin-top:8px;font-size:13px"><a href="/diagnostica" target="_blank">Diagnostica</a></div>
    <div style="margin-top:12px;color:#888;font-size:12px">Tocca per chiudere</div>
  </div>
</div>
<div id="qrov" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.85);z-index:9;align-items:center;justify-content:center">
  <img src="/qr.svg" alt="QR" style="width:min(80vw,80vh);height:min(80vw,80vh);background:#fff">
</div>

<script>
const DEF = { font:"Arial, sans-serif", size:48, fg:"#ffffff", bg:"#000000", lh:1.3, ls:0 };
let S = Object.assign({}, DEF);
try { Object.assign(S, JSON.parse(localStorage.getItem("pptNotesStyle") || "{}")); } catch(e) {}
const $ = id => document.getElementById(id);
const notes = $("notes");
const lsGet = (k, d) => { try { const v = localStorage.getItem(k); return v === null ? d : v; } catch(e) { return d; } };
const lsSet = (k, v) => { try { localStorage.setItem(k, v); } catch(e) {} };

function apply() {
  const r = document.documentElement.style;
  r.setProperty("--fg", S.fg); r.setProperty("--bg", S.bg);
  notes.style.fontFamily = S.font;
  notes.style.fontSize = S.size + "px";
  notes.style.lineHeight = S.lh;
  notes.style.letterSpacing = S.ls + "px";
  $("font").value = S.font; $("size").value = S.size;
  $("fg").value = S.fg; $("bg").value = S.bg;
  $("lh").value = S.lh; $("ls").value = S.ls;
  $("lhv").textContent = Number(S.lh).toFixed(2);
  $("lsv").textContent = Number(S.ls).toFixed(1) + "px";
  lsSet("pptNotesStyle", JSON.stringify(S));
}
function bind(id, key, conv) {
  $(id).addEventListener("input", e => { S[key] = conv ? conv(e.target.value) : e.target.value; apply(); });
}
bind("font","font"); bind("size","size",v=>Math.max(10, Math.min(300, parseInt(v)||DEF.size)));
bind("fg","fg"); bind("bg","bg"); bind("lh","lh",parseFloat); bind("ls","ls",parseFloat);
$("sizeUp").onclick = () => { S.size = Math.min(300, S.size + 4); apply(); };
$("sizeDown").onclick = () => { S.size = Math.max(10, S.size - 4); apply(); };
$("reset").onclick = () => { S = Object.assign({}, DEF); apply(); };
$("full").onclick = () => {
  if (document.fullscreenElement) document.exitFullscreen();
  else document.documentElement.requestFullscreen().catch(()=>{});
};
$("qrb").onclick = () => { $("qrov").style.display = "flex"; };
$("infob").onclick = () => { $("infoov").style.display = "flex"; };
$("infoov").onclick = () => { $("infoov").style.display = "none"; };
$("qrov").onclick = () => { $("qrov").style.display = "none"; };
$("hide").onclick = () => document.body.classList.add("hidebar");
$("showbar").onclick = () => document.body.classList.remove("hidebar");
apply();

/* ---- stato, scelta presentazione ---- */
let st = { shows: [], pin_required: false, control: true };
let selId = lsGet("pptSel", "");
let lastKey = null, optKey = "";

function currentShow() {
  if (!st.shows.length) return null;
  return st.shows.find(s => s.id === selId) || st.shows[0];
}
$("show").onchange = e => { selId = e.target.value; lsSet("pptSel", selId); lastKey = null; render(); };

function render() {
  const cur = currentShow();
  $("dot").className = cur ? "on" : "";
  // elenco presentazioni (visibile solo se ce n'e' piu' di una)
  $("showlbl").style.display = st.shows.length > 1 ? "" : "none";
  const key = st.shows.map(s => s.id + "|" + s.title).join("§");
  const sel = $("show");
  if (key !== optKey) {
    optKey = key;
    sel.innerHTML = "";
    st.shows.forEach(s => { const o = document.createElement("option"); o.value = s.id; sel.appendChild(o); });
  }
  st.shows.forEach((s, i) => { sel.options[i].textContent = s.title + " — " + s.index + "/" + s.total; });
  if (cur) sel.value = cur.id;

  if (!cur) {
    notes.textContent = "In attesa della presentazione…";
    notes.classList.add("wait");
    $("pos").textContent = ""; $("ctlpos").textContent = "–";
    lastKey = null;
    updateSlides(null);
    return;
  }
  notes.classList.remove("wait");
  $("pos").textContent = cur.title + "  —  slide " + cur.index + " / " + cur.total;
  $("ctlpos").textContent = cur.index + " / " + cur.total;
  $("gonum").max = cur.total;
  notes.textContent = cur.notes && cur.notes.trim() ? cur.notes : "(nessuna nota per questa slide)";
  const k = cur.id + "#" + cur.index;
  if (k !== lastKey) { $("main").scrollTop = 0; lastKey = k; }
  updateSlides(cur);
}

/* ---- anteprima slide in onda e successiva ---- */
function setSlide(img, ph, on, id, index, title, emptyText) {
  if (!on) return;
  if (!id || !index) { img.style.display = "none"; ph.textContent = emptyText; ph.style.display = "flex"; return; }
  const src = "/slide/" + encodeURIComponent(id) + "/" + index + ".png?v=" + encodeURIComponent(title);
  if (img.getAttribute("data-src") !== src) { img.setAttribute("data-src", src); img.src = src; }
  ph.style.display = "none"; img.style.display = "block";
}
function updateSlides(cur) {
  const wait = "In attesa della presentazione…";
  setSlide($("imgcur"), $("phcur"), curOn, cur ? cur.id : "", cur ? cur.index : 0, cur ? cur.title : "", wait);
  const hasNext = !!cur && cur.index < cur.total;
  setSlide($("imgnext"), $("phnext"), nextOn, hasNext ? cur.id : "", hasNext ? cur.index + 1 : 0,
           cur ? cur.title : "", cur ? "Ultima slide" : wait);
}
["imgcur", "imgnext"].forEach(id => {
  $(id).onerror = () => {
    const ph = $(id === "imgcur" ? "phcur" : "phnext");
    const src = $(id).getAttribute("data-src");
    $(id).style.display = "none"; $(id).removeAttribute("data-src");
    ph.textContent = "Anteprima non disponibile"; ph.style.display = "flex";
    // chiede al programma il motivo e lo mostra al posto dell'anteprima
    if (src) fetch(src).then(r => r.text()).then(t => {
      if (t) ph.textContent = "Anteprima non disponibile: " + t.slice(0, 220);
    }).catch(() => {});
  };
});

/* ---- comandi ---- */
let pin = lsGet("pptPin", "");
let msgTimer = null;
function msg(t) {
  $("ctlmsg").textContent = t;
  clearTimeout(msgTimer);
  if (t) msgTimer = setTimeout(() => { $("ctlmsg").textContent = ""; }, 3000);
}
async function sendCmd(action, n, retry) {
  const cur = currentShow();
  if (!cur) { msg("Nessuna presentazione"); return; }
  try {
    const r = await fetch("/cmd", { method:"POST", headers:{ "Content-Type":"application/json" },
      body: JSON.stringify({ id: cur.id, action: action, n: n, pin: pin }) });
    if (r.status === 403) {
      let err = ""; try { err = (await r.json()).error; } catch(e) {}
      if (err === "pin") {
        if (retry) { msg("PIN errato"); pin = ""; lsSet("pptPin", ""); return; }
        const p = prompt("PIN per i comandi:");
        if (p === null) return;
        pin = p; lsSet("pptPin", pin);
        return sendCmd(action, n, true);
      }
      msg("Comandi disabilitati"); return;
    }
    if (!r.ok) msg("Errore comando");
  } catch(e) { msg("Connessione persa"); }
}
$("next").onclick = () => sendCmd("next");
$("prev").onclick = () => sendCmd("prev");
$("go").onclick = () => { const n = parseInt($("gonum").value); if (n > 0) sendCmd("goto", n); };
$("gonum").addEventListener("keydown", e => { if (e.key === "Enter") $("go").click(); });
document.addEventListener("keydown", e => {
  if (!ctlOn || e.target.tagName === "INPUT" || e.target.tagName === "SELECT") return;
  if (["ArrowRight","ArrowDown","PageDown"," "].includes(e.key)) { e.preventDefault(); sendCmd("next"); }
  else if (["ArrowLeft","ArrowUp","PageUp"].includes(e.key)) { e.preventDefault(); sendCmd("prev"); }
});
/* ---- cosa mostrare: ogni dispositivo ricorda le sue scelte ---- */
let styleOn = lsGet("pptStyle", "1") === "1";
let ctlOn = lsGet("pptCtl", "1") === "1";
let curOn = lsGet("pptCur", "0") === "1";
let nextOn = lsGet("pptNext", "0") === "1";
function applyCtl() {
  const allowed = st.control !== false;
  const b = document.body.classList;
  b.toggle("nostyle", !styleOn);
  b.toggle("noctl", !ctlOn || !allowed);
  b.toggle("showcur", curOn);
  b.toggle("shownext", nextOn);
  $("ctlb").style.display = allowed ? "" : "none";
  $("styleb").classList.toggle("on", styleOn);
  $("ctlb").classList.toggle("on", ctlOn);
  $("curb").classList.toggle("on", curOn);
  $("nextb").classList.toggle("on", nextOn);
}
$("styleb").onclick = () => { styleOn = !styleOn; lsSet("pptStyle", styleOn ? "1" : "0"); applyCtl(); };
$("ctlb").onclick = () => { ctlOn = !ctlOn; lsSet("pptCtl", ctlOn ? "1" : "0"); applyCtl(); };
$("curb").onclick = () => { curOn = !curOn; lsSet("pptCur", curOn ? "1" : "0"); applyCtl(); render(); };
$("nextb").onclick = () => { nextOn = !nextOn; lsSet("pptNext", nextOn ? "1" : "0"); applyCtl(); render(); };

/* ---- schermo sempre acceso (serve HTTPS o app; sul sito http della LAN molti browser lo negano) ---- */
let wl = null;
async function keepAwake() {
  try {
    if ("wakeLock" in navigator && !wl) {
      wl = await navigator.wakeLock.request("screen");
      wl.addEventListener("release", () => { wl = null; });
    }
  } catch(e) { wl = null; }
}
document.addEventListener("visibilitychange", () => { if (document.visibilityState === "visible") keepAwake(); });
keepAwake();

/* ---- collegamento in tempo reale ---- */
function connect() {
  const es = new EventSource("/events");
  es.onmessage = e => { try { st = JSON.parse(e.data); applyCtl(); render(); } catch(err) {} };
  es.onerror = () => { $("dot").className = ""; };
}
applyCtl();
connect();
</script>
</body>
</html>
"""


def main():
    global BASE_URL, PIN, CONTROL, DEMO_MODE
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    ap = argparse.ArgumentParser()
    ap.add_argument("--port", type=int, default=8080)
    ap.add_argument("--interval", type=float, default=0.1, help="secondi tra una lettura e l'altra")
    ap.add_argument("--pin", default="", help="PIN richiesto per i comandi avanti/indietro")
    ap.add_argument("--no-control", action="store_true", help="disabilita i comandi dal telefono")
    ap.add_argument("--demo", action="store_true", help="simula presentazioni (senza PowerPoint)")
    ap.add_argument("--no-browser", action="store_true", help="non aprire la pagina QR sul PC")
    args = ap.parse_args()
    PIN = str(args.pin or "")
    CONTROL = not args.no_control
    DEMO_MODE = bool(args.demo)

    if not args.demo:
        try:
            import win32com.client  # noqa: F401
        except Exception:
            print("ERRORE: serve Windows con pywin32 (e PowerPoint installato).")
            print("Per provare senza PowerPoint usa --demo.")
            input("Premi Invio per chiudere...")
            return

    srv = None
    port = args.port
    for p in range(args.port, args.port + 20):
        try:
            srv = ThreadingHTTPServer(("0.0.0.0", p), Handler)
            port = p
            break
        except OSError:
            continue
    if srv is None:
        print("ERRORE: nessuna porta libera tra %d e %d." % (args.port, args.port + 19))
        input("Premi Invio per chiudere...")
        return
    srv.daemon_threads = True

    BASE_URL = "http://%s:%d/" % (primary_ip(), port)
    target = poll_demo if args.demo else poll_powerpoint
    threading.Thread(target=target, args=(args.interval,), daemon=True).start()

    try:
        import ctypes
        ctypes.windll.kernel32.SetConsoleTitleW("%s %s" % (APP_NAME, APP_VERSION))
    except Exception:
        pass
    print("=" * 60)
    print(" %s %s - note del relatore sempre con te" % (APP_NAME, APP_VERSION))
    print("=" * 60)
    print(" Indirizzo per smartphone e altri PC della rete:")
    print("   " + BASE_URL)
    others = ["http://%s:%d/" % (i, port) for i in local_ips() if ("http://%s:%d/" % (i, port)) != BASE_URL]
    if others:
        print(" Altri indirizzi possibili (se ci sono piu' schede di rete):")
        for o in others:
            print("   " + o)
    print()
    try:
        print(qr_console(BASE_URL))
    except Exception as e:
        print(" (QR in console non disponibile: %s)" % e)
    print()
    print(" Se il QR non si legge bene: pagina grande su http://localhost:%d/qr" % port)
    if not CONTROL:
        print(" Comandi avanti/indietro: DISABILITATI")
    elif PIN:
        print(" Comandi avanti/indietro: attivi, protetti da PIN")
    else:
        print(" Comandi avanti/indietro: attivi, SENZA PIN (chiunque sulla rete puo' usarli;")
        print("   per proteggerli avvia con --pin 1234)")
    print(" Alla prima esecuzione Windows chiede di consentire l'accesso alla rete:")
    print(" scegli 'Reti private' (e 'Reti pubbliche' se il Wi-Fi e' visto come pubblico).")
    print(" Registro errori: " + log_path())
    print(" Diagnostica: http://localhost:%d/diagnostica" % port)
    print(" Chiudi questa finestra per fermare il programma.")
    print()
    print(" reMNote - Ideato e sviluppato da www.andreabattaglioli.it")
    print(" Powered by www.mywebtvlive.com - Made in Bolognina")

    log("Avvio %s %s - indirizzo %s" % (APP_NAME, APP_VERSION, BASE_URL))
    if not args.no_browser:
        try:
            webbrowser.open("http://localhost:%d/qr" % port)
        except Exception:
            pass
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    try:
        main()
    except Exception:
        import traceback
        traceback.print_exc()
        input("Errore. Premi Invio per chiudere...")
