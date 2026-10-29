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
import socket
import sys
import threading
import time
import webbrowser
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler

APP_NAME = "reMNote"
APP_VERSION = "1.0"
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
            if int(app.SlideShowWindows.Count) > 0:
                shows = read_shows(app)
        except Exception:
            # PowerPoint chiuso o COM momentaneamente occupato: ricollega al giro dopo
            app = None
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
.hd{display:flex;align-items:center;gap:12px;font-size:34px;font-weight:700;color:#14434d}
.hd img{width:56px;height:56px}.cr{font-size:14px;color:#444;line-height:1.5}
.cr a{color:#14434d}.cr img{vertical-align:middle;margin-right:4px}</style></head>
<body><div class="hd"><img src="/static/remnote-192.png" alt="">reMNote</div>
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
  #infobox img.logo { width:96px; height:96px; }
  #infobox h2 { margin:6px 0 0; color:#14434d; }
  #infobox .ver { color:#666; font-size:13px; margin-bottom:12px; }
  #infobox a { color:#14434d; }
  #infobox .mwtl img { vertical-align:middle; margin-right:4px; }
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
  <button id="full">Schermo intero</button>
  <button id="ctlb">Comandi</button>
  <button id="qrb">QR</button>
  <button id="infob">Info</button>
  <button id="hide">Nascondi barra</button>
  <div id="info"><span id="pos"></span><span id="dot" title="Presentazione"></span></div>
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
    <h2>reMNote</h2>
    <div class="ver">Versione __VERSION__</div>
    <div>__CREDITS__</div>
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
    return;
  }
  notes.classList.remove("wait");
  $("pos").textContent = cur.title + "  —  slide " + cur.index + " / " + cur.total;
  $("ctlpos").textContent = cur.index + " / " + cur.total;
  $("gonum").max = cur.total;
  notes.textContent = cur.notes && cur.notes.trim() ? cur.notes : "(nessuna nota per questa slide)";
  const k = cur.id + "#" + cur.index;
  if (k !== lastKey) { $("main").scrollTop = 0; lastKey = k; }
}

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
let ctlOn = lsGet("pptCtl", "1") === "1";
function applyCtl() {
  const allowed = st.control !== false;
  document.body.classList.toggle("noctl", !ctlOn || !allowed);
  $("ctlb").style.display = allowed ? "" : "none";
  $("ctlb").textContent = ctlOn ? "Nascondi comandi" : "Comandi";
}
$("ctlb").onclick = () => { ctlOn = !ctlOn; lsSet("pptCtl", ctlOn ? "1" : "0"); applyCtl(); };

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
    global BASE_URL, PIN, CONTROL
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
    print(" Chiudi questa finestra per fermare il programma.")
    print()
    print(" reMNote - Ideato e sviluppato da www.andreabattaglioli.it")
    print(" Powered by www.mywebtvlive.com - Made in Bolognina")

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
