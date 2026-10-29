#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DeckWhisperer - Remote notes & control for PowerPoint
Legge le note del relatore da PowerPoint (in presentazione) e le pubblica su una
pagina web raggiungibile dalla rete locale, in tempo reale.

Funzioni:
  - piu' presentazioni aperte contemporaneamente: ciascun dispositivo sceglie quale seguire
  - comandi per avanzare/tornare/andare a una slide dal telefono (PIN opzionale)
  - schermo sempre acceso (dove il browser lo consente)

Uso:
    py deckwhisperer.py                  (Windows, con PowerPoint installato)
    py deckwhisperer.py --pin 1234       (i comandi richiedono il PIN)
    py deckwhisperer.py --no-control     (disabilita del tutto i comandi)
    py deckwhisperer.py --demo           (prova senza PowerPoint)
    py deckwhisperer.py --port 8080
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
import zlib
import webbrowser
from collections import OrderedDict, deque
from html import escape as html_escape
from xml.sax.saxutils import escape
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler

APP_NAME = "MEFISTOFELE DeckWhisperer"
APP_SUBTITLE = "Remote notes & control for PowerPoint"
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
    "deckwhisperer-512.png": "image/png", "deckwhisperer-192.png": "image/png", "deckwhisperer-64.png": "image/png",
    "deckwhisperer-32.png": "image/png", "apple-touch-icon.png": "image/png", "mwtl-logo.png": "image/png",
    "deckwhisperer.ico": "image/x-icon",
}


def resource_path(*parts):
    """Percorso di un file incluso nel programma (anche dentro l'exe PyInstaller)."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, *parts)


CREDITS_HTML = (
    '<b>MEFISTOFELE DeckWhisperer</b> – Remote notes &amp; control for PowerPoint<br>Ideato e sviluppato da '
    '<a href="https://www.andreabattaglioli.it" target="_blank" rel="noopener">www.andreabattaglioli.it</a> · '
    '<a class="mwtl" href="https://www.mywebtvlive.com" target="_blank" rel="noopener">'
    '<img src="/static/mwtl-logo.png" alt="" width="18" height="18">MyWebTvLive.com</a> · Made in Bolognina'
)

UI_KEYS = ("bar", "style", "cur", "next", "goto", "ctl")
# barra in alto, stile testo, slide in onda, slide successiva, "vai alla slide", avanti/indietro
UI = {k: True for k in UI_KEYS}              # cosa possono vedere i telefoni (scelto dal PC)
NIC_IP = ""                                  # scheda di rete scelta per il QR ("" = automatica)
PORT = 8080              # cosa possono vedere i telefoni (scelto dal PC)
NO_CONTROL = False                           # --no-control: comandi spenti, qualunque sia la scelta
SETTINGS_LOCK = threading.Lock()
STATE = {"pin_required": False, "control": True, "ui": dict(UI), "shows": []}
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
    new = {"pin_required": bool(PIN), "control": CONTROL, "ui": ui_state(), "shows": shows}
    with COND:
        if new != STATE:
            STATE.clear()
            STATE.update(new)
            VERSION += 1
            COND.notify_all()


def ui_state():
    """Cosa mostrare sui telefoni: la scelta fatta dal PC, con i comandi spenti se serve."""
    ui = dict(UI)
    ui["ctl"] = bool(UI["ctl"]) and not NO_CONTROL
    ui["goto"] = bool(UI["goto"]) and not NO_CONTROL
    return ui


def refresh_state():
    """Rimanda lo stato ai telefoni dopo un cambio di impostazioni (senza aspettare PowerPoint)."""
    global CONTROL
    CONTROL = ui_state()["ctl"]
    with COND:
        shows = list(STATE.get("shows", []))
    set_state(shows)


def list_nics():
    """Schede di rete con un indirizzo IPv4 utilizzabile: [{'name':..., 'ip':...}]."""
    out, seen = [], set()
    try:
        import psutil
        stats = psutil.net_if_stats()
        for name, addrs in psutil.net_if_addrs().items():
            st = stats.get(name)
            if st is not None and not st.isup:
                continue
            for a in addrs:
                ip = a.address
                if a.family == socket.AF_INET and not ip.startswith(("127.", "169.254.")) and ip not in seen:
                    seen.add(ip)
                    out.append({"name": str(name), "ip": ip})
    except Exception:
        pass
    if not out:
        out = [{"name": "Scheda di rete", "ip": ip} for ip in local_ips()]
    return out


def current_ip():
    """IP scritto nel QR: la scheda scelta dal PC, se c'e' ancora, altrimenti quella automatica."""
    if NIC_IP and any(n["ip"] == NIC_IP for n in list_nics()):
        return NIC_IP
    return primary_ip()


def update_base_url():
    global BASE_URL
    BASE_URL = "http://%s:%d/" % (current_ip(), PORT)


def settings_paths(name="DeckWhisperer.json"):
    """Dove salvare le impostazioni: accanto all'exe, poi nella cartella temporanea."""
    return [os.path.join(os.path.dirname(p), name) for p in log_paths()]


def load_settings():
    """Legge le impostazioni salvate (se ci sono). Ritorna il PIN salvato."""
    for pth in settings_paths() + settings_paths("reMNote.json"):   # l'ultimo e' il nome usato dalla 2.0
        try:
            with open(pth, encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            continue
        global NIC_IP
        ui = data.get("ui", {})
        for k in UI_KEYS:
            if isinstance(ui.get(k), bool):
                UI[k] = ui[k]
        if isinstance(data.get("nic"), str):
            NIC_IP = data["nic"][:45]
        return str(data.get("pin", ""))[:20]
    return ""


def save_settings():
    data = json.dumps({"ui": UI, "pin": PIN, "nic": NIC_IP}, ensure_ascii=False, indent=1)
    for pth in settings_paths():
        try:
            with open(pth, "w", encoding="utf-8") as f:
                f.write(data)
            return pth
        except Exception:
            continue
    log("Impostazioni non salvate: nessuna cartella scrivibile")
    return ""


def apply_settings(body):
    """Applica le scelte arrivate dalla pagina di avvio. Ritorna False se non valide."""
    global PIN, NIC_IP
    ui = body.get("ui")
    if not isinstance(ui, dict) or any(k in ui and not isinstance(ui[k], bool) for k in UI_KEYS):
        return False
    pin = str(body.get("pin", PIN))
    if len(pin) > 20 or any(c in pin for c in "\r\n\t"):
        return False
    nic = body.get("nic", NIC_IP)
    if not isinstance(nic, str) or (nic and nic not in [n["ip"] for n in list_nics()]):
        return False
    with SETTINGS_LOCK:
        for k in UI_KEYS:
            if k in ui:
                UI[k] = ui[k]
        PIN = pin.strip()
        NIC_IP = nic
        save_settings()
    update_base_url()
    refresh_state()
    return True


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
    """Id stabile della finestra di presentazione, fatto solo di cifre/lettere.
    Con alcune installazioni di PowerPoint/pywin32 HWND arriva come metodo e non come
    numero: va chiamato, altrimenti l'id diventa "<bound method ...>" (uguale per tutte)."""
    try:
        h = win.HWND
        if callable(h):
            h = h()
        h = int(h)
        if h:
            return str(abs(h))
    except Exception:
        pass
    try:
        name = str(win.Presentation.FullName)
    except Exception:
        name = ""
    if name:
        return "p%08x" % (zlib.crc32(name.encode("utf-8", "replace")) & 0xFFFFFFFF)
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
    return os.path.join(base, "DeckWhisperer.log")


def log_paths():
    """Dove prova a scrivere il registro: accanto all'exe, poi nella cartella temporanea."""
    return [log_path(), os.path.join(tempfile.gettempdir(), "DeckWhisperer.log")]


def log(msg):
    """Scrive un messaggio in console, in memoria e nel file DeckWhisperer.log (ogni messaggio una volta sola)."""
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
        "DeckWhisperer %s" % APP_VERSION,
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
<meta name="viewport" content="width=device-width, initial-scale=1"><title>MEFISTOFELE DeckWhisperer - diagnostica</title>
<link rel="icon" type="image/png" href="/static/deckwhisperer-64.png">
<style>body{font:14px Arial,sans-serif;margin:16px;background:#fff;color:#111}
pre{background:#f3f3f3;padding:10px;border-radius:6px;white-space:pre-wrap;overflow-wrap:anywhere}
button{font:15px Arial,sans-serif;padding:8px 14px;margin:4px 6px 4px 0;cursor:pointer}</style></head>
<body><h2>MEFISTOFELE DeckWhisperer - diagnostica</h2>
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
            path = os.path.join(tempfile.gettempdir(), "DeckWhisperer_%d_%d.png" % (os.getpid(), job["index"]))
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


START_PAGE = """<!DOCTYPE html><html lang="it"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1"><title>__NAME__ - avvio</title>
<link rel="icon" type="image/png" href="/static/deckwhisperer-64.png">
<style>
:root{--bg:#f4f6f7;--card:#fff;--ink:#14262b;--mut:#5b6b70;--acc:#14434d;--ok:#1b7f4a;--bd:#d9e0e2}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font-family:Segoe UI,Arial,sans-serif}
.wrap{max-width:980px;margin:0 auto;padding:20px 16px 28px}
header{display:flex;align-items:center;gap:14px;margin-bottom:18px;flex-wrap:wrap}
header img{width:64px;height:64px}h1{font-size:26px;margin:0;letter-spacing:.02em}header small{color:var(--mut);display:block;margin-top:2px}
.guide{margin-left:auto;font-size:16px;padding:10px 18px;text-decoration:none;background:var(--acc);color:#fff;border-radius:8px}
.grid{display:grid;grid-template-columns:minmax(260px,380px) 1fr;gap:18px}
@media(max-width:760px){.grid{grid-template-columns:1fr}}
.card{background:var(--card);border:1px solid var(--bd);border-radius:12px;padding:18px}
.card h2{font-size:17px;margin:0 0 12px}
.qr{text-align:center}.qr img{width:100%;max-width:320px;height:auto}
.url{font-size:20px;font-weight:700;word-break:break-all;margin:8px 0 2px}.hint{color:var(--mut);font-size:14px}
.nic{margin-top:14px;text-align:left}.nic label{font-weight:700;font-size:14px;display:block;margin-bottom:5px}
.nic select{width:100%;font-size:15px;padding:8px;border:1px solid var(--bd);border-radius:8px;background:#fff}
.opt{display:flex;gap:12px;align-items:flex-start;padding:11px 4px;border-bottom:1px solid var(--bd);cursor:pointer}
.opt:last-of-type{border-bottom:0}.opt input{width:22px;height:22px;margin-top:2px;accent-color:var(--acc);flex:none}
.opt b{display:block;font-size:16px}.opt span{color:var(--mut);font-size:14px}
.pin{display:flex;gap:10px;align-items:center;margin-top:14px;flex-wrap:wrap}
.pin input{font-size:18px;padding:8px 10px;width:150px;border:1px solid var(--bd);border-radius:8px}
button{font-size:15px;padding:9px 14px;border:0;border-radius:8px;background:var(--acc);color:#fff;cursor:pointer}
#saved{color:var(--ok);font-size:14px;min-height:20px;margin-top:10px}
#shows{margin:0;padding-left:18px;color:var(--ink)}#shows li{margin:2px 0}
.cr{font-size:13px;color:var(--mut);line-height:1.6;text-align:center;margin-top:22px}.cr a{color:var(--acc)}
.cr img{vertical-align:middle;margin-right:4px}.links{margin-top:10px;font-size:14px}.links a{color:var(--acc)}
</style></head><body><div class="wrap">
<header><img src="/static/deckwhisperer-192.png" alt="__NAME__"><div><h1>__NAME__</h1>
<small>Versione __VERSION__ &middot; Remote notes &amp; control for PowerPoint</small></div>
<a class="guide" href="/guida" target="_blank">&#128214; Guida</a></header>
<div class="grid">
<div class="card qr"><h2>1. Inquadra il QR</h2><img id="qrimg" src="/qr.svg" alt="QR">
<div class="url" id="url">__URL__</div><div class="hint">Telefono collegato alla stessa rete Wi-Fi del PC</div>
<div class="nic"><label for="nic">Scheda di rete del QR</label><select id="nic"></select>
<div class="hint" style="margin-top:5px">Scegli la scheda a cui sono collegati i telefoni (es. quella della rete di produzione).</div></div></div>
<div class="card"><h2>2. Cosa vuoi far vedere sui telefoni</h2>
<label class="opt"><input type="checkbox" id="u_bar"><div><b>Barra dei comandi in alto</b><span>Pulsanti colorati, QR, info. Spenta: restano note, scelta presentazione e la striscia con lucchetto</span></div></label>
<label class="opt"><input type="checkbox" id="u_style"><div><b>Comandi stile testo</b><span>Font, dimensione, colori, interlinea e spaziatura</span></div></label>
<label class="opt"><input type="checkbox" id="u_cur"><div><b>Slide in onda</b><span>Immagine della slide che il pubblico sta vedendo</span></div></label>
<label class="opt"><input type="checkbox" id="u_next"><div><b>Slide successiva</b><span>Immagine della slide che arriva dopo</span></div></label>
<label class="opt"><input type="checkbox" id="u_goto"><div><b>Vai direttamente alla slide</b><span>Campo numero e pulsante Vai per saltare a una slide</span></div></label>
<label class="opt"><input type="checkbox" id="u_ctl"><div><b>Avanti / Indietro</b><span>Pulsanti per far avanzare e tornare indietro le slide</span></div></label>
<div class="pin"><label for="pin"><b>PIN per i comandi</b> (facoltativo)</label>
<input id="pin" maxlength="20" inputmode="numeric" autocomplete="off" placeholder="nessuno"><button id="savepin">Salva PIN</button></div>
<div id="saved"></div>
<h2 style="margin-top:14px">Presentazioni in corso</h2><ul id="shows"><li>nessuna: apri una presentazione in PowerPoint</li></ul>
<div class="links"><a href="/" target="_blank">Apri la pagina delle note su questo PC</a> &middot; <a href="/diagnostica" target="_blank">Diagnostica</a></div>
</div></div>
<div class="cr">__CREDITS__</div></div>
<script>
const $ = id => document.getElementById(id), K = ["bar","style","cur","next","goto","ctl"];
let timer = null;
function flash(t){ $("saved").textContent = t; clearTimeout(timer); timer = setTimeout(() => { $("saved").textContent = ""; }, 2500); }
async function save(){
  const ui = {}; K.forEach(k => { ui[k] = $("u_" + k).checked; });
  try {
    const r = await fetch("/settings", { method:"POST", headers:{ "Content-Type":"application/json" },
      body: JSON.stringify({ ui: ui, pin: $("pin").value.trim(), nic: $("nic").value }) });
    flash(r.ok ? "Salvato: i telefoni collegati si aggiornano subito" : "Impostazioni non valide");
    if (r.ok) load(false);
  } catch(e) { flash("Programma non raggiungibile"); }
}
function load(first){
  fetch("/settings.json").then(r => r.json()).then(d => {
    if (first) {
      K.forEach(k => { $("u_" + k).checked = d.ui[k] !== false; });
      $("pin").value = d.pin || "";
      const sel = $("nic"); sel.innerHTML = "";
      const auto = document.createElement("option"); auto.value = "";
      auto.textContent = "Automatica (consigliata)" + (d.auto_ip ? " - " + d.auto_ip : ""); sel.appendChild(auto);
      d.nics.forEach(n => { const o = document.createElement("option"); o.value = n.ip; o.textContent = n.name + " - " + n.ip; sel.appendChild(o); });
      sel.value = d.nic || "";
      if (sel.value !== (d.nic || "")) sel.value = "";
      $("u_ctl").disabled = !!d.no_control; $("u_goto").disabled = !!d.no_control;
      if (d.no_control) $("u_ctl").parentNode.querySelector("span").textContent = "Disattivati: il programma è stato avviato con --no-control";
    }
    $("url").textContent = d.url; $("qrimg").src = "/qr.svg?t=" + Date.now();
  }).catch(() => flash("Programma non raggiungibile"));
}
K.forEach(k => { $("u_" + k).onchange = save; });
$("nic").onchange = save;
$("savepin").onclick = save;
$("pin").addEventListener("keydown", e => { if (e.key === "Enter") save(); });
load(true);
function shows(){
  fetch("/state").then(r => r.json()).then(s => {
    const ul = $("shows"); ul.innerHTML = "";
    if (!s.shows.length) { ul.innerHTML = "<li>nessuna: apri una presentazione in PowerPoint</li>"; return; }
    s.shows.forEach(x => { const li = document.createElement("li"); li.textContent = x.title + " - slide " + x.index + " / " + x.total; ul.appendChild(li); });
  }).catch(() => {});
}
shows(); setInterval(shows, 2000);
</script></body></html>"""


GUIDE_FILES = {"it": "ISTRUZIONI.md", "en": "INSTRUCTIONS.md"}


def guide_path(lang):
    name = GUIDE_FILES.get(lang, GUIDE_FILES["it"])
    for cand in (resource_path("docs", name), os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "docs", name)):
        if os.path.exists(cand):
            return cand
    return ""


def _inline(t):
    t = html_escape(t, quote=False)
    t = re.sub(r"`([^`]+)`", r"<code>\1</code>", t)
    t = re.sub(r"\*\*([^*]+)\*\*", r"<b>\1</b>", t)
    t = re.sub(r"\*([^*]+)\*", r"<i>\1</i>", t)

    def link(m):
        url = m.group(2)
        if url.endswith("INSTRUCTIONS.md"):
            url = "/guida?lang=en"
        elif url.endswith("ISTRUZIONI.md"):
            url = "/guida?lang=it"
        if not url.startswith(("http://", "https://", "/")):
            return m.group(1)
        return '<a href="%s" target="_blank" rel="noopener">%s</a>' % (url, m.group(1))
    return re.sub(r"\[([^\]]+)\]\(([^)\s]+)\)", link, t)


def md_to_html(md):
    """Converte il Markdown delle guide (titoli, elenchi, tabelle, grassetto, codice, link) in HTML."""
    out, para, items, rows = [], [], [], []

    def flush():
        nonlocal para, items, rows
        if para:
            out.append("<p>%s</p>" % _inline(" ".join(para)))
        if items:
            tag = "ol" if items[0][0] else "ul"
            out.append("<%s>%s</%s>" % (tag, "".join("<li>%s</li>" % _inline(t) for _, t in items), tag))
        if rows:
            head, body = rows[0], rows[2:]
            out.append("<table><tr>%s</tr>%s</table>" % (
                "".join("<th>%s</th>" % _inline(c) for c in head),
                "".join("<tr>%s</tr>" % "".join("<td>%s</td>" % _inline(c) for c in r) for r in body)))
        para, items, rows = [], [], []

    for line in md.splitlines():
        st = line.strip()
        if not st:
            flush()
        elif st.startswith("#"):
            flush()
            n = len(st) - len(st.lstrip("#"))
            out.append("<h%d>%s</h%d>" % (n, _inline(st[n:].strip()), n))
        elif st == "---":
            flush()
            out.append("<hr>")
        elif st.startswith("|"):
            if para or items:
                flush()
            rows.append([c.strip() for c in st.strip("|").split("|")])
        elif re.match(r"^(\d+\.|-)\s+", st):
            if para or rows:
                flush()
            items.append((st[0].isdigit(), re.sub(r"^(\d+\.|-)\s+", "", st)))
        elif items and line.startswith(" "):      # riga di continuazione di una voce dell'elenco
            items[-1] = (items[-1][0], items[-1][1] + " " + st)
        else:
            if items or rows:
                flush()
            para.append(st)
    flush()
    return "\n".join(out)


GUIDE_PAGE = """<!DOCTYPE html><html lang="__LANG__"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1"><title>__NAME__ - __TITLE__</title>
<link rel="icon" type="image/png" href="/static/deckwhisperer-64.png">
<style>body{margin:0;background:#f4f6f7;color:#14262b;font:16px/1.6 Segoe UI,Arial,sans-serif}
.wrap{max-width:860px;margin:0 auto;padding:18px 16px 40px;background:#fff;min-height:100vh}
.top{display:flex;gap:10px;align-items:center;flex-wrap:wrap;border-bottom:1px solid #d9e0e2;padding-bottom:10px;margin-bottom:6px}
.top img{width:44px;height:44px}.top a{margin-left:auto}.sw a{margin-left:10px;font-weight:700}
h1{font-size:26px}h2{font-size:21px;margin-top:28px}h3{font-size:17px}
table{border-collapse:collapse;width:100%;margin:10px 0}th,td{border:1px solid #d9e0e2;padding:7px 10px;text-align:left;vertical-align:top}
th{background:#eef3f4}code{background:#eef3f4;padding:1px 5px;border-radius:4px}a{color:#14434d}</style></head><body><div class="wrap">
<div class="top"><img src="/static/deckwhisperer-64.png" alt=""><b>__NAME__</b>
<span class="sw"><a href="/guida?lang=it">Italiano</a><a href="/guida?lang=en">English</a></span></div>
__BODY__</div></body></html>"""


def guide_page(lang):
    lang = lang if lang in GUIDE_FILES else "it"
    pth = guide_path(lang)
    if pth:
        with open(pth, encoding="utf-8") as f:
            body = md_to_html(f.read())
    else:
        body = "<p>Guida non trovata / Guide not found.</p>"
    return (GUIDE_PAGE.replace("__LANG__", lang).replace("__NAME__", html_escape(APP_NAME))
            .replace("__TITLE__", "Guida" if lang == "it" else "Guide").replace("__BODY__", body))


# -------------------------------------------------------------------- Server
class Handler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"

    def log_message(self, *a):
        pass

    def _is_local(self):
        """True se la richiesta arriva dal PC stesso (le impostazioni si cambiano solo da qui)."""
        ip = self.client_address[0]
        return ip in ("127.0.0.1", "::1") or ip.startswith("::ffff:127.") or ip in local_ips()

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
        if path == "/settings":
            if not self._is_local():
                self._json(403, {"ok": False, "error": "solo dal PC"})
                return
            try:
                ok = apply_settings(json.loads(raw.decode("utf-8") or "{}"))
            except Exception:
                ok = False
            self._json(200 if ok else 400, {"ok": ok})
            return
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
            name = "deckwhisperer.ico" if path == "/favicon.ico" else path[len("/static/"):]
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
        elif path in ("/start", "/qr"):
            if not self._is_local():
                self._send(403, "text/plain; charset=utf-8", "Pagina di avvio: si apre solo dal PC dove gira DeckWhisperer".encode("utf-8"))
                return
            self._send(200, "text/html; charset=utf-8", START_PAGE.replace("__URL__", BASE_URL).replace("__NAME__", html_escape(APP_NAME))
                       .replace("__VERSION__", APP_VERSION).replace("__CREDITS__", CREDITS_HTML).encode("utf-8"))
        elif path == "/guida":
            q = self.path.split("?", 1)[1] if "?" in self.path else ""
            lang = "en" if "lang=en" in q else "it"
            self._send(200, "text/html; charset=utf-8", guide_page(lang).encode("utf-8"))
        elif path == "/settings.json":
            if not self._is_local():
                self._send(403, "text/plain; charset=utf-8", b"Solo dal PC")
                return
            self._send(200, "application/json; charset=utf-8",
                       json.dumps({"ui": UI, "pin": PIN, "no_control": NO_CONTROL, "url": BASE_URL,
                                   "nics": list_nics(), "nic": NIC_IP, "auto_ip": primary_ip()}).encode("utf-8"))
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
<title>MEFISTOFELE DeckWhisperer</title>
<link rel="icon" type="image/png" sizes="64x64" href="/static/deckwhisperer-64.png">
<link rel="icon" type="image/png" sizes="32x32" href="/static/deckwhisperer-32.png">
<link rel="apple-touch-icon" href="/static/apple-touch-icon.png">
<meta name="theme-color" content="#14434d">
<style>
  :root { --fg:#ffffff; --bg:#000000; }
  * { box-sizing: border-box; }
  html, body { margin:0; height:100%; background:var(--bg); color:var(--fg); }
  body { display:flex; flex-direction:column; font-family: Arial, sans-serif; }
  #bar { max-height:46vh; overflow-y:auto; display:flex; flex-wrap:wrap; gap:8px 14px; align-items:center; padding:6px 10px;
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
  #ctl .goto { display:flex; gap:6px; align-items:center; }
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
  /* ---- pulsanti colorati: ogni funzione ha il suo colore; pieno = acceso, bordo = spento ---- */
  .chips { display:flex; flex-wrap:wrap; gap:8px; width:100%; }
  #bar .tg, #bar .act { min-height:44px; padding:8px 16px; border-radius:24px; font:600 15px Arial, sans-serif;
        cursor:pointer; border:2px solid var(--c); background:transparent; color:var(--c); }
  #bar .tg::before { content:"\25CB\00a0"; }
  #bar .tg.on { background:var(--c); color:#fff; }
  #bar .tg.on::before { content:"\25CF\00a0"; }
  #bar .act { background:var(--c); color:#fff; }
  .c-style { --c:#a47bff; } .c-ctl { --c:#1fae62; } .c-cur { --c:#ff7a45; } .c-next { --c:#3b9cff; }
  .c-lock { --c:#e0534f; } .c-full { --c:#12a8a8; } .c-qr { --c:#7d8b92; } .c-info { --c:#7d8b92; } .c-hide { --c:#c8a23a; }
  .sec { font:700 11px Arial, sans-serif; letter-spacing:.08em; text-transform:uppercase; color:#9fb0b6; width:100%; }
  #ctl button { min-height:64px; }
  #prev { background:#d98a14 !important; border-color:#f2b24d !important; color:#1b1200 !important; }
  #next { background:#1fae62 !important; border-color:#6fe0a2 !important; color:#05200f !important; }
  #go { background:#3b9cff !important; border-color:#9fcdff !important; color:#fff !important; min-height:36px; }
  #ctl .big:active { filter:brightness(1.25); transform:scale(.98); }
  body.nobar #bar > *:not(#showlbl):not(#info) { display:none !important; }
  body.nobar #showbar { display:none !important; }
  body.hidebar #bar { display:none; }
  body.noctl #ctl { display:none; }
  /* gruppo stile testo: caratteri +50% e pulsanti piu' grandi */
  #bar .grp label { font-size:20px; }
  #bar .grp select, #bar .grp input[type=number], #bar .grp button { font-size:20px; min-height:44px; }
  #bar .grp input[type=number] { width:84px; }
  #bar .grp input[type=color] { width:54px; height:40px; }
  #bar .grp input[type=range] { width:130px; }
  #bar .grp #lhv, #bar .grp #lsv { font-size:20px; }
  #bar .grp #sizeDown, #bar .grp #sizeUp { min-width:58px; font-weight:700; }
  /* striscia in alto quando la barra e' chiusa (o non abilitata) e quando lo schermo e' bloccato */
  #strip { display:none; flex-wrap:wrap; gap:6px; align-items:center; padding:6px 8px; background:#222;
           border-bottom:1px solid #444; position:relative; z-index:7; }
  body.hidebar #strip, body.nobar #strip, body.locked #strip { display:flex; }
  #strip button { min-height:44px; padding:6px 8px; border-radius:22px; border:2px solid var(--c); background:transparent;
                  color:var(--c); font:600 14px Arial, sans-serif; cursor:pointer; white-space:nowrap; flex:1 1 auto; }
  #strip button.on { background:var(--c); color:#fff; }
  #sbar { --c:#c8a23a; } #sfull { --c:#12a8a8; } #sctl { --c:#1fae62; } #slock { --c:#ff5a5a; }
  body.nobar #sbar { display:none; } body.nostripctl #sctl { display:none; }
  #sunlock { display:none; --c:#ff5a5a; position:relative; overflow:hidden; flex:1; user-select:none;
             -webkit-user-select:none; -webkit-touch-callout:none; touch-action:none; }
  #sunlock .prog { position:absolute; left:0; top:0; bottom:0; width:0; background:rgba(255,90,90,.55); }
  #sunlock.holding .prog { width:100%; transition:width var(--hold,5s) linear; }
  #sunlock span { position:relative; }
  body.locked #sunlock { display:block; }
  body.locked #strip > button:not(#sunlock) { display:none; }
  body.locked #bar, body.locked #sbar { display:none; }
  #lockov { display:none; position:fixed; left:0; right:0; top:58px; bottom:0; z-index:6; touch-action:none; background:transparent; }
  body.locked #lockov { display:block; }
  body.nogoto .goto { display:none; }
</style>
</head>
<body>
<div id="strip">
  <button id="sbar" title="Riapri la barra">&#9776; Barra</button>
  <button id="sfull" title="Schermo intero">&#9974; Schermo</button>
  <button id="sctl" class="on" title="Mostra/nascondi Avanti e Indietro">&#9664;&#9654; Comandi</button>
  <button id="slock" title="Blocca lo schermo">&#128274; Blocca</button>
  <button id="sunlock" title="Tieni premuto 5 secondi per sbloccare"><span class="prog"></span><span>&#128275; Tieni premuto 5 secondi per sbloccare</span></button>
</div>
<div id="lockov"></div>
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
    <input type="number" id="size" min="10" max="300" step="2" inputmode="numeric">
    <button id="sizeDown" title="Riduci">a&minus;</button>
    <button id="sizeUp" title="Aumenta">A+</button>
  </label>
  <label>Testo <input type="color" id="fg"></label>
  <label>Sfondo <input type="color" id="bg"></label>
  <label>Interlinea <input type="range" id="lh" min="0.9" max="3" step="0.05"> <span id="lhv"></span></label>
  <label>Spaziatura <input type="range" id="ls" min="-2" max="20" step="0.5"> <span id="lsv"></span></label>
  <button id="reset">Ripristina</button>
  </span>
  <div class="sec">Cosa mostrare</div>
  <div class="chips">
    <button id="styleb" class="tg c-style" title="Mostra/nascondi i comandi del testo">Stile testo</button>
    <button id="ctlb" class="tg c-ctl" title="Mostra/nascondi i pulsanti per avanzare le slide">Comandi slide</button>
    <button id="curb" class="tg c-cur" title="Mostra/nascondi la slide in onda">Slide in onda</button>
    <button id="nextb" class="tg c-next" title="Mostra/nascondi la slide successiva">Slide successiva</button>
  </div>
  <div class="sec">Strumenti</div>
  <div class="chips">
    <button id="full" class="act c-full">Schermo intero</button>
    <button id="qrb" class="act c-qr">QR</button>
    <button id="infob" class="act c-info">Info</button>
    <button id="lockb" class="act c-lock">&#128274; Blocca schermo</button>
    <button id="hide" class="act c-hide">Nascondi barra</button>
  </div>
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
    <span class="goto"><input type="number" id="gonum" min="1" inputmode="numeric"> <button id="go">Vai</button></span>
    <span id="ctlmsg"></span>
  </div>
  <button id="next" class="big">Avanti ▶</button>
</div>
<div id="infoov">
  <div id="infobox">
    <img class="logo" src="/static/deckwhisperer-192.png" alt="DeckWhisperer">
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
$("hide").onclick = () => { document.body.classList.add("hidebar"); applyCtl(); };
$("sbar").onclick = () => { document.body.classList.remove("hidebar"); applyCtl(); };
$("sfull").onclick = () => $("full").onclick();
document.addEventListener("fullscreenchange", () => { $("sfull").classList.toggle("on", !!document.fullscreenElement); });
apply();

/* ---- stato, scelta presentazione ---- */
let st = { shows: [], pin_required: false, control: true, ui: { bar: true, style: true, cur: true, next: true, goto: true, ctl: true } };
const uiOn = k => !st.ui || st.ui[k] !== false;   // scelto dal PC nella pagina di avvio
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
  setSlide($("imgcur"), $("phcur"), curOn && uiOn("cur"), cur ? cur.id : "", cur ? cur.index : 0, cur ? cur.title : "", wait);
  const hasNext = !!cur && cur.index < cur.total;
  setSlide($("imgnext"), $("phnext"), nextOn && uiOn("next"), hasNext ? cur.id : "", hasNext ? cur.index + 1 : 0,
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
  if (!ctlOn || !uiOn("ctl") || st.control === false || e.target.tagName === "INPUT" || e.target.tagName === "SELECT") return;
  if (["ArrowRight","ArrowDown","PageDown"," "].includes(e.key)) { e.preventDefault(); sendCmd("next"); }
  else if (["ArrowLeft","ArrowUp","PageUp"].includes(e.key)) { e.preventDefault(); sendCmd("prev"); }
});
/* ---- cosa mostrare: ogni dispositivo ricorda le sue scelte ---- */
let styleOn = lsGet("pptStyle", "1") === "1";
let ctlOn = lsGet("pptCtl", "1") === "1";
let curOn = lsGet("pptCur", "1") === "1";
let nextOn = lsGet("pptNext", "1") === "1";
function applyCtl() {
  const allowed = st.control !== false && uiOn("ctl");
  const b = document.body.classList;
  b.toggle("nostyle", !styleOn || !uiOn("style"));
  b.toggle("noctl", !ctlOn || !allowed);
  b.toggle("nobar", !uiOn("bar"));
  b.toggle("nogoto", !uiOn("goto"));
  b.toggle("nostripctl", !allowed);
  $("sctl").classList.toggle("on", ctlOn);
  b.toggle("showcur", curOn && uiOn("cur"));
  b.toggle("shownext", nextOn && uiOn("next"));
  $("styleb").style.display = uiOn("style") ? "" : "none";
  $("ctlb").style.display = allowed ? "" : "none";
  $("curb").style.display = uiOn("cur") ? "" : "none";
  $("nextb").style.display = uiOn("next") ? "" : "none";
  $("styleb").classList.toggle("on", styleOn);
  $("ctlb").classList.toggle("on", ctlOn);
  $("curb").classList.toggle("on", curOn);
  $("nextb").classList.toggle("on", nextOn);
}
$("styleb").onclick = () => { styleOn = !styleOn; lsSet("pptStyle", styleOn ? "1" : "0"); applyCtl(); };
$("ctlb").onclick = () => { ctlOn = !ctlOn; lsSet("pptCtl", ctlOn ? "1" : "0"); applyCtl(); };
$("curb").onclick = () => { curOn = !curOn; lsSet("pptCur", curOn ? "1" : "0"); applyCtl(); render(); };
$("nextb").onclick = () => { nextOn = !nextOn; lsSet("pptNext", nextOn ? "1" : "0"); applyCtl(); render(); };

/* ---- schermo sempre acceso: Wake Lock (se il browser lo consente) + video silenzioso in loop (funziona anche su http) ---- */
const NOSLEEP_WEBM = "data:video/webm;base64,GkXfo59ChoEBQveBAULygQRC84EIQoKEd2VibUKHgQJChYECGFOAZwEAAAAAAAJKEU2bdLpNu4tTq4QVSalmU6yBoU27i1OrhBZUrmtTrIHYTbuMU6uEElTDZ1OsggEeTbuMU6uEHFO7a1OsggI07AEAAAAAAABZAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAVSalmsirXsYMPQkBNgI1MYXZmNjAuMTYuMTAwV0GNTGF2ZjYwLjE2LjEwMESJiECPQAAAAAAAFlSua8GuAQAAAAAAADjXgQFzxYg9xTkZCNtYGJyBACK1nIN1bmSIgQCGhVZfVlA4g4EBI+ODhAvrwgDgibCBQLqBQJqBAhJUw2f8c3OgY8CAZ8iaRaOHRU5DT0RFUkSHjUxhdmY2MC4xNi4xMDBzc9ZjwItjxYg9xTkZCNtYGGfIoUWjh0VOQ09ERVJEh5RMYXZjNjAuMzEuMTAyIGxpYnZweGfIoUWjiERVUkFUSU9ORIeTMDA6MDA6MDEuMDAwMDAwMDAwAB9DtnVAj+eBAKOqgQAAgPACAJ0BKkAAQAAARwiFhYiFhIgCAgAGcDxCYAqyIPcwAP7/q1CAo5aBAMgA0QEAARAQABgAGFgv9AAIjoAAo5aBAZAA0QEAARAQABgAGFgv9AAIjoAAo5aBAlgA0QEAARAQABgAGFgv9AAIjoAAo5aBAyAA0QEAARAQABgAGFgv9AAIjoAAHFO7a5G7j7OBALeK94EB8YIBn/CBAw==";
const NOSLEEP_MP4 = "data:video/mp4;base64,AAAAIGZ0eXBpc29tAAACAGlzb21pc28yYXZjMW1wNDEAAAM2bW9vdgAAAGxtdmhkAAAAAAAAAAAAAAAAAAAD6AAAA+gAAQAAAQAAAAAAAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAAAmB0cmFrAAAAXHRraGQAAAADAAAAAAAAAAAAAAABAAAAAAAAA+gAAAAAAAAAAAAAAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAAAAABAAAAAAEAAAABAAAAAAAAkZWR0cwAAABxlbHN0AAAAAAAAAAEAAAPoAAAAAAABAAAAAAHYbWRpYQAAACBtZGhkAAAAAAAAAAAAAAAAAAAoAAAAKABVxAAAAAAALWhkbHIAAAAAAAAAAHZpZGUAAAAAAAAAAAAAAABWaWRlb0hhbmRsZXIAAAABg21pbmYAAAAUdm1oZAAAAAEAAAAAAAAAAAAAACRkaW5mAAAAHGRyZWYAAAAAAAAAAQAAAAx1cmwgAAAAAQAAAUNzdGJsAAAAt3N0c2QAAAAAAAAAAQAAAKdhdmMxAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAAAAEAAQABIAAAASAAAAAAAAAABFUxhdmM2MC4zMS4xMDIgbGlieDI2NAAAAAAAAAAAAAAAGP//AAAALWF2Y0MBQsAK/+EAFmdCwAraEJsBEAAAAwAQAAADAKDxImoBAARozg/IAAAAEHBhc3AAAAABAAAAAQAAABRidHJ0AAAAAAAAFcgAABXIAAAAGHN0dHMAAAAAAAAAAQAAAAUAAAgAAAAAFHN0c3MAAAAAAAAAAQAAAAEAAAAcc3RzYwAAAAAAAAABAAAAAQAAAAUAAAABAAAAKHN0c3oAAAAAAAAAAAAAAAUAAAKRAAAACgAAAAoAAAAKAAAACgAAABRzdGNvAAAAAAAAAAEAAANmAAAAYnVkdGEAAABabWV0YQAAAAAAAAAhaGRscgAAAAAAAAAAbWRpcmFwcGwAAAAAAAAAAAAAAAAtaWxzdAAAACWpdG9vAAAAHWRhdGEAAAABAAAAAExhdmY2MC4xNi4xMDAAAAAIZnJlZQAAAsFtZGF0AAACbwYF//9r3EXpvebZSLeWLNgg2SPu73gyNjQgLSBjb3JlIDE2NCByMzEwOCAzMWUxOWY5IC0gSC4yNjQvTVBFRy00IEFWQyBjb2RlYyAtIENvcHlsZWZ0IDIwMDMtMjAyMyAtIGh0dHA6Ly93d3cudmlkZW9sYW4ub3JnL3gyNjQuaHRtbCAtIG9wdGlvbnM6IGNhYmFjPTAgcmVmPTEgZGVibG9jaz0xOjA6MCBhbmFseXNlPTB4MToweDExMSBtZT1oZXggc3VibWU9MiBwc3k9MSBwc3lfcmQ9MS4wMDowLjAwIG1peGVkX3JlZj0wIG1lX3JhbmdlPTE2IGNocm9tYV9tZT0xIHRyZWxsaXM9MCA4eDhkY3Q9MCBjcW09MCBkZWFkem9uZT0yMSwxMSBmYXN0X3Bza2lwPTEgY2hyb21hX3FwX29mZnNldD0wIHRocmVhZHM9MiBsb29rYWhlYWRfdGhyZWFkcz0xIHNsaWNlZF90aHJlYWRzPTAgbnI9MCBkZWNpbWF0ZT0xIGludGVybGFjZWQ9MCBibHVyYXlfY29tcGF0PTAgY29uc3RyYWluZWRfaW50cmE9MCBiZnJhbWVzPTAgd2VpZ2h0cD0wIGtleWludD0yNTAga2V5aW50X21pbj01IHNjZW5lY3V0PTQwIGludHJhX3JlZnJlc2g9MCByY19sb29rYWhlYWQ9MTAgcmM9Y3JmIG1idHJlZT0xIGNyZj0yMy4wIHFjb21wPTAuNjAgcXBtaW49MCBxcG1heD02OSBxcHN0ZXA9NCBpcF9yYXRpbz0xLjQwIGFxPTE6MS4wMACAAAAAGmWIhAR///8PRQABKr8nJyddddddddddddeAAAAABkGaIBHwjAAAAAZBmkAR8IwAAAAGQZpgEPCMAAAABkGagD/CMA==";
let wl = null, nsVideo = null;
async function keepAwake() {
  try {
    if ("wakeLock" in navigator && !wl) {
      wl = await navigator.wakeLock.request("screen");
      wl.addEventListener("release", () => { wl = null; });
    }
  } catch(e) { wl = null; }
  try {
    if (!nsVideo) {
      nsVideo = document.createElement("video");
      nsVideo.setAttribute("playsinline", ""); nsVideo.setAttribute("muted", "");
      nsVideo.setAttribute("title", "DeckWhisperer"); nsVideo.muted = true; nsVideo.loop = true;
      nsVideo.style.cssText = "position:fixed;left:0;top:0;width:1px;height:1px;opacity:.01;pointer-events:none";
      nsVideo.src = (nsVideo.canPlayType && nsVideo.canPlayType("video/webm")) ? NOSLEEP_WEBM : NOSLEEP_MP4;
      if (document.body.appendChild) document.body.appendChild(nsVideo);
    }
    if (nsVideo.paused && nsVideo.play) { const pr = nsVideo.play(); if (pr && pr.catch) pr.catch(() => {}); }
  } catch(e) {}
}
// i browser vogliono un tocco dell'utente per far partire il video: si riprova a ogni tocco finche' parte
["pointerdown", "touchstart", "click"].forEach(t => document.addEventListener(t, () => { if (!nsVideo || nsVideo.paused) keepAwake(); }));
document.addEventListener("visibilitychange", () => { if (document.visibilityState === "visible") keepAwake(); });
keepAwake();

/* ---- blocco schermo: lucchetto, si sblocca tenendo premuto 5 secondi ---- */
let LOCK_HOLD_MS = 5000, lockTimer = null;
function setLocked(on) {
  document.body.classList.toggle("locked", on);
  if (on) { $("infoov").style.display = "none"; $("qrov").style.display = "none"; }
  endHold();
}
function startHold(e) {
  if (e && e.preventDefault) e.preventDefault();
  endHold();
  $("sunlock").style.setProperty("--hold", (LOCK_HOLD_MS / 1000) + "s");
  $("sunlock").classList.add("holding");
  lockTimer = setTimeout(() => { lockTimer = null; setLocked(false); }, LOCK_HOLD_MS);
}
function endHold() {
  if (lockTimer) { clearTimeout(lockTimer); lockTimer = null; }
  $("sunlock").classList.remove("holding");
}
$("slock").onclick = () => setLocked(true);
$("lockb").onclick = () => setLocked(true);
$("sunlock").onpointerdown = startHold;
$("sunlock").onpointerup = endHold;
$("sunlock").onpointerleave = endHold;
$("sunlock").onpointercancel = endHold;
$("sunlock").oncontextmenu = e => { if (e && e.preventDefault) e.preventDefault(); return false; };
$("lockov").oncontextmenu = e => { if (e && e.preventDefault) e.preventDefault(); return false; };
$("sctl").onclick = () => $("ctlb").onclick();

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
    global BASE_URL, PIN, CONTROL, DEMO_MODE, NO_CONTROL, PORT
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
    ap.add_argument("--no-browser", action="store_true", help="non aprire la pagina di avvio sul PC")
    args = ap.parse_args()
    NO_CONTROL = bool(args.no_control)
    DEMO_MODE = bool(args.demo)
    saved_pin = load_settings()                  # scelte fatte l'ultima volta dalla pagina di avvio
    PIN = str(args.pin or saved_pin or "")
    CONTROL = ui_state()["ctl"]

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

    PORT = port
    update_base_url()
    target = poll_demo if args.demo else poll_powerpoint
    threading.Thread(target=target, args=(args.interval,), daemon=True).start()

    try:
        import ctypes
        ctypes.windll.kernel32.SetConsoleTitleW("%s %s" % (APP_NAME, APP_VERSION))
    except Exception:
        pass
    print("=" * 60)
    print(" %s %s - Remote notes & control for PowerPoint" % (APP_NAME, APP_VERSION))
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
    print(" Pagina di avvio (QR e scelta di cosa mostrare sui telefoni): http://localhost:%d/start" % port)
    if not CONTROL:
        print(" Comandi avanti/indietro: DISABILITATI")
    elif PIN:
        print(" Comandi avanti/indietro: attivi, protetti da PIN")
    else:
        print(" Comandi avanti/indietro: attivi, SENZA PIN (chiunque sulla rete puo' usarli;")
        print("   per proteggerli imposta un PIN nella pagina di avvio)")
    print(" Alla prima esecuzione Windows chiede di consentire l'accesso alla rete:")
    print(" scegli 'Reti private' (e 'Reti pubbliche' se il Wi-Fi e' visto come pubblico).")
    print(" Registro errori: " + log_path())
    print(" Diagnostica: http://localhost:%d/diagnostica" % port)
    print(" Chiudi questa finestra per fermare il programma.")
    print()
    print(" DeckWhisperer - Ideato e sviluppato da www.andreabattaglioli.it")
    print(" Powered by www.mywebtvlive.com - Made in Bolognina")

    log("Avvio %s %s - indirizzo %s" % (APP_NAME, APP_VERSION, BASE_URL))
    if not args.no_browser:
        try:
            webbrowser.open("http://localhost:%d/start" % port)
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
