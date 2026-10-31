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
import base64
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
    '<b>MEFISTOFELE DeckWhisperer</b> – Remote notes &amp; control for PowerPoint<br><span id="credby">Created and developed by</span> '
    '<a href="https://www.andreabattaglioli.it" target="_blank" rel="noopener">www.andreabattaglioli.it</a> · '
    '<a class="mwtl" href="https://www.mywebtvlive.com" target="_blank" rel="noopener">'
    '<img src="static/mwtl-logo.png" alt="" width="18" height="18">MyWebTvLive.com</a> · Made in Bolognina'
)

UI_KEYS = ("bar", "style", "cur", "next", "goto", "ctl", "lock")
# barra degli strumenti, stile testo, slide in onda, slide successiva, vai alla slide, navigazione, blocca schermo remoto
UI = {k: (k != "lock") for k in UI_KEYS}              # cosa possono vedere i telefoni (scelto dal PC)
NIC_IP = ""                                  # scheda di rete scelta per il QR ("" = automatica)
LANG = "en"                                  # lingua predefinita dei telefoni (inglese)
PORT = 8080              # cosa possono vedere i telefoni (scelto dal PC)
NO_CONTROL = False                           # --no-control: comandi spenti, qualunque sia la scelta
SETTINGS_LOCK = threading.Lock()
FOCUS_KEEP = True                             # dopo ogni anteprima rida il fuoco a PowerPoint se era suo
def load_i18n():
    """Traduzioni dell'interfaccia (assets/i18n.json): inglese, italiano, francese, spagnolo, arabo, cinese."""
    try:
        with open(resource_path("assets", "i18n.json"), encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {"names": {"en": "English"}, "rtl": [], "langs": {"en": {}}}


I18N = load_i18n()
LANG_CODES = tuple(I18N["names"])
I18N_JSON = json.dumps(I18N, ensure_ascii=False).replace("</", "<\\/")

STATE = {"pin_required": False, "control": True, "ui": dict(UI), "lang": LANG, "shows": []}
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
    new = {"pin_required": bool(PIN), "control": CONTROL, "ui": ui_state(), "lang": LANG, "shows": shows}
    with COND:
        if new != STATE:
            STATE.clear()
            STATE.update(new)
            VERSION += 1
            COND.notify_all()


def ui_state():
    """Cosa mostrare sui telefoni: la scelta fatta dal PC, con i comandi spenti se serve."""
    ui = dict(UI)
    stop = NO_CONTROL or bool(UI["lock"])        # --no-control o schermo remoto bloccato: nessun comando
    ui["ctl"] = bool(UI["ctl"]) and not stop
    ui["goto"] = bool(UI["goto"]) and not stop
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


def settings_paths(name="mefistofele.json"):
    """Dove salvare le impostazioni: accanto all'exe, poi nella cartella temporanea."""
    return [os.path.join(os.path.dirname(p), name) for p in log_paths()]


def load_settings():
    """Legge le impostazioni salvate (se ci sono). Ritorna il PIN salvato."""
    for pth in settings_paths() + settings_paths("DeckWhisperer.json") + settings_paths("reMNote.json"):   # i vecchi nomi (2.x)
        try:
            with open(pth, encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            continue
        global NIC_IP, LANG, FOCUS_KEEP
        ui = data.get("ui", {})
        if isinstance(data.get("focus"), bool):
            FOCUS_KEEP = data["focus"]
        for k in UI_KEYS:
            if isinstance(ui.get(k), bool):
                UI[k] = ui[k]
        if isinstance(data.get("nic"), str):
            NIC_IP = data["nic"][:45]
        if data.get("lang") in LANG_CODES:
            LANG = data["lang"]
        return str(data.get("pin", ""))[:20]
    return ""


def save_settings():
    data = json.dumps({"ui": UI, "pin": PIN, "nic": NIC_IP, "lang": LANG, "focus": FOCUS_KEEP}, ensure_ascii=False, indent=1)
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
    global PIN, NIC_IP, LANG, FOCUS_KEEP
    ui = body.get("ui")
    if not isinstance(ui, dict) or any(k in ui and not isinstance(ui[k], bool) for k in UI_KEYS):
        return False
    pin = str(body.get("pin", PIN))
    if len(pin) > 20 or any(c in pin for c in "\r\n\t"):
        return False
    nic = body.get("nic", NIC_IP)
    if not isinstance(nic, str) or (nic and nic not in [n["ip"] for n in list_nics()]):
        return False
    lang = body.get("lang", LANG)
    if lang not in LANG_CODES:
        return False
    focus = body.get("focus", FOCUS_KEEP)
    if not isinstance(focus, bool):
        return False
    with SETTINGS_LOCK:
        for k in UI_KEYS:
            if k in ui:
                UI[k] = ui[k]
        PIN = pin.strip()
        NIC_IP = nic
        LANG = lang
        FOCUS_KEEP = focus
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
    return os.path.join(base, "mefistofele.log")


def log_paths():
    """Dove prova a scrivere il registro: accanto all'exe, poi nella cartella temporanea."""
    return [log_path(), os.path.join(tempfile.gettempdir(), "mefistofele.log")]


def log(msg):
    """Scrive un messaggio in console, in memoria e nel file mefistofele.log (ogni messaggio una volta sola)."""
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


LOG_EV_COUNT = [0]
CMD_MAX_AGE = 15.0     # un comando piu' vecchio di cosi' non ha piu' senso: si scarta, ma si scrive nel registro


def log_ev(msg):
    """Come log(), ma senza togliere i messaggi uguali (servono i tempi): al massimo 300 righe per avvio."""
    if LOG_EV_COUNT[0] >= 300:
        return
    LOG_EV_COUNT[0] += 1
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


def drain_commands(runner):
    """Esegue i comandi in coda. Avanti/indietro consecutivi sulla stessa presentazione
    vengono sommati in un solo passaggio; si scarta (e si registra) solo cio' che e' molto vecchio."""
    cmds = []
    while True:
        try:
            cmds.append(CMDQ.get_nowait())
        except queue.Empty:
            break
    if not cmds:
        return 0
    t0 = time.time()
    batch = []
    for c in cmds:
        age = t0 - c.get("t", t0)
        if age > CMD_MAX_AGE:
            log_ev("Comando scartato: troppo vecchio (%.1f s): %s" % (age, c.get("action")))
            continue
        act = c.get("action")
        if act in ("next", "prev"):
            step = 1 if act == "next" else -1
            if batch and batch[-1].get("action") == "step" and batch[-1]["id"] == c.get("id"):
                batch[-1]["n"] += step
                continue
            batch.append({"id": c.get("id"), "action": "step", "n": step})
        else:
            batch.append(c)
    done = 0
    for b in batch:
        try:
            if b["action"] == "step":
                act = "next" if b["n"] > 0 else "prev"
                for _ in range(abs(b["n"])):
                    runner({"id": b["id"], "action": act})
                    done += 1
            else:
                runner(b)
                done += 1
        except Exception as e:
            log_ev("Comando %s non eseguito: %s: %s" % (b.get("action"), type(e).__name__, e))
    dt = time.time() - t0
    if len(cmds) > 1 or dt > 0.5:
        log_ev("Comandi: %d ricevuti, %d passaggi eseguiti in %.2f s" % (len(cmds), done, dt))
    return len(cmds)


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
<link rel="icon" type="image/png" href="static/deckwhisperer-64.png">
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
        job["error"] = "presentation not found (id %s)" % job["id"]
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
                        errs.append("image file not created")
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


def _fg_window():
    import ctypes
    return int(ctypes.windll.user32.GetForegroundWindow() or 0)


def _window_pid(hwnd):
    import ctypes
    from ctypes import wintypes
    pid = wintypes.DWORD(0)
    ctypes.windll.user32.GetWindowThreadProcessId(wintypes.HWND(hwnd), ctypes.byref(pid))
    return int(pid.value)


def _set_fg_window(hwnd):
    """SetForegroundWindow, con il trucco dell'accodamento dell'input se Windows rifiuta."""
    import ctypes
    u = ctypes.windll.user32
    if u.SetForegroundWindow(hwnd):
        return True
    try:
        fg = u.GetForegroundWindow()
        t1 = u.GetWindowThreadProcessId(fg, None)
        t2 = ctypes.windll.kernel32.GetCurrentThreadId()
        if t1 and t1 != t2:
            u.AttachThreadInput(t2, t1, True)
            ok = u.SetForegroundWindow(hwnd)
            u.AttachThreadInput(t2, t1, False)
            return bool(ok)
    except Exception:
        pass
    return False


def powerpoint_pids(app):
    """Processi delle finestre di presentazione (di solito uno solo)."""
    pids = set()
    for i in range(1, int(app.SlideShowWindows.Count) + 1):
        try:
            h = app.SlideShowWindows(i).HWND
            if callable(h):
                h = h()
            pids.add(_window_pid(int(h)))
        except Exception:
            pass
    pids.discard(0)
    return pids


def focus_snapshot(app):
    """(finestra in primo piano, True se appartiene a PowerPoint). Mai un errore: senza Windows ritorna (0, False)."""
    if not FOCUS_KEEP or sys.platform != "win32":
        return 0, False
    try:
        fg = _fg_window()
        return fg, bool(fg) and _window_pid(fg) in powerpoint_pids(app)
    except Exception:
        return 0, False


def focus_restore(app, before, get_fg=None, pid_of=None, set_fg=None, pp_pids=None, wait=0.1):
    """Dopo un'operazione che puo' togliere il fuoco (export di un'anteprima): se prima era di PowerPoint e ora
    non lo e' piu', lo rida. Non interviene se ora e' in primo piano un'altra applicazione (scelta dell'utente)."""
    hwnd, was_pp = before
    if not (hwnd and was_pp):
        return False
    get_fg = get_fg or _fg_window
    pid_of = pid_of or _window_pid
    set_fg = set_fg or _set_fg_window
    try:
        pids = pp_pids if pp_pids is not None else powerpoint_pids(app)
        own = os.getpid()
        for attempt in range(2):
            now = get_fg()
            if now == hwnd:
                if attempt == 0:
                    if wait:
                        time.sleep(wait)      # PowerPoint a volte toglie il fuoco un attimo dopo l'export
                    continue
                return False
            if now == 0 or pid_of(now) in pids or pid_of(now) == own:
                ok = set_fg(hwnd)
                log("Fuoco restituito a PowerPoint dopo un'anteprima" + ("" if ok else " (Windows ha rifiutato)"))
                return bool(ok)
            return False                  # un'altra applicazione ha il fuoco: non e' compito nostro
    except Exception:
        return False
    return False


def process_image_jobs(app):
    while True:
        drain_commands(lambda c: run_command(app, c))   # i comandi hanno la precedenza sulle anteprime
        try:
            job = IMGQ.get_nowait()
        except queue.Empty:
            break
        if time.time() - job["t"] > 8.0:
            job["event"].set()   # nessuno la sta piu' aspettando
            continue
        try:
            t_exp = time.time()
            fbefore = focus_snapshot(app)
            export_slide(app, job)
            focus_restore(app, fbefore)
            if time.time() - t_exp > 1.5:
                log_ev("Anteprima lenta: slide %s in %.1f s" % (job["index"], time.time() - t_exp))
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
            drain_commands(lambda c: run_command(app, c))
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


# ---------------------------------------------------------------- PowerPoint per Mac (AppleScript)
IS_MAC = sys.platform == "darwin"
OSA_SEP_F, OSA_SEP_R = "\x1f", "\x1e"

OSA_COMMON = """
on notesOf(sl)
  tell application "Microsoft PowerPoint"
    try
      set np to notes page of sl
      set fallback to ""
      repeat with s in shapes of np
        try
          set t to content of text range of text frame of s
          if (length of t) > (length of fallback) then set fallback to t
          set k to (placeholder type of placeholder format of s) as string
          if k contains "body" then return t
        end try
      end repeat
      return fallback
    end try
  end tell
  return ""
end notesOf
"""

OSA_READ = OSA_COMMON + """
on run argv
  set FS to ASCII character 31
  set RS to ASCII character 30
  set out to ""
  tell application "Microsoft PowerPoint"
    set n to count of slide show windows
    repeat with i from 1 to n
      try
        set w to slide show window i
        set v to slide show view of w
        set idx to current show position of v
        set p to presentation of w
        set tot to count of slides of p
        set cur to my notesOf(slide idx of p)
        set nxt to ""
        if idx < tot then set nxt to my notesOf(slide (idx + 1) of p)
        set out to out & (full name of p) & FS & (name of p) & FS & idx & FS & tot & FS & cur & FS & nxt & RS
      end try
    end repeat
  end tell
  return out
end run
"""

OSA_CMD = """
on run argv
  set k to (item 1 of argv) as integer
  set act to item 2 of argv
  tell application "Microsoft PowerPoint"
    set v to slide show view of slide show window k
    if act is "next" then
      go to next slide v
    else if act is "prev" then
      go to previous slide v
    else if act is "first" then
      go to first slide v
    else if act is "last" then
      go to last slide v
    else if act is "goto" then
      go to slide v number ((item 3 of argv) as integer)
    end if
  end tell
  return "ok"
end run
"""

OSA_EXPORT = """
on run argv
  set k to (item 1 of argv) as integer
  set n to (item 2 of argv) as integer
  tell application "Microsoft PowerPoint"
    set p to presentation of slide show window k
    save slide n of p in (POSIX file (item 3 of argv)) as save as PNG
  end tell
  return "ok"
end run
"""


def osa(script, *args, timeout=6):
    """Esegue uno script AppleScript con osascript e ritorna il testo restituito."""
    import subprocess
    r = subprocess.run(["osascript", "-"] + [str(a) for a in args], input=script,
                       capture_output=True, text=True, timeout=timeout)
    if r.returncode != 0:
        raise RuntimeError((r.stderr or "osascript error").strip()[:300])
    return r.stdout


def mac_read_shows():
    """Come read_shows() ma con AppleScript. Ritorna (lista presentazioni, nomi completi nell'ordine delle finestre)."""
    out = osa(OSA_READ)
    shows, order = [], []
    for rec in out.split(OSA_SEP_R):
        f = rec.strip("\n").split(OSA_SEP_F)
        if len(f) < 6:
            continue
        full, name, idx, tot, cur, nxt = f[:6]
        norm = lambda t: t.replace("\r\n", "\n").replace("\r", "\n").replace("\x0b", "\n")
        shows.append({"id": "p%08x" % (zlib.crc32(full.encode("utf-8", "replace")) & 0xFFFFFFFF),
                      "title": name, "index": int(float(idx)), "total": int(float(tot)),
                      "notes": norm(cur), "next_notes": norm(nxt)})
        order.append(shows[-1]["id"])
    return shows, order


def mac_window_of(order, wid):
    for i, x in enumerate(order, 1):
        if not wid or x == wid:
            return i
    return 0


def mac_run_command(order, cmd):
    k = mac_window_of(order, cmd.get("id"))
    if not k:
        return False
    act = cmd.get("action")
    n = ""
    if act == "goto":
        n = int(cmd.get("n"))
    osa(OSA_CMD, k, act, n)
    return True


def mac_export(order, job):
    key = (job["id"], job["index"])
    data = cache_get(key)
    job["error"] = ""
    if data is None:
        k = mac_window_of(order, job["id"]) if job["id"] in order else 0
        if not k:
            job["error"] = "presentation not found (id %s)" % job["id"]
        else:
            path = os.path.join(tempfile.gettempdir(), "DeckWhisperer_%d_%d.png" % (os.getpid(), job["index"]))
            try:
                osa(OSA_EXPORT, k, job["index"], path, timeout=8)
                try:   # riduce a 960 px di larghezza (sips e' incluso in macOS)
                    import subprocess
                    subprocess.run(["sips", "--resampleWidth", "960", path], capture_output=True, timeout=5)
                except Exception:
                    pass
                if os.path.exists(path) and os.path.getsize(path) > 0:
                    with open(path, "rb") as f:
                        data = f.read()
                else:
                    job["error"] = "image file not created"
            except Exception as e:
                job["error"] = "export failed (%s: %s)" % (type(e).__name__, e)
            finally:
                try:
                    os.remove(path)
                except OSError:
                    pass
        if data:
            cache_put(key, data)
    job["result"] = data
    if not data:
        log("Anteprima slide %s/%s: %s" % (job["id"], job["index"], job["error"]))
    job["event"].set()


def poll_powerpoint_mac(interval):
    interval = max(interval, 0.4)       # osascript costa piu' di una chiamata COM
    order = []
    while True:
        shows = []
        try:
            drain_commands(lambda c: mac_run_command(order, c))
            while True:
                drain_commands(lambda c: mac_run_command(order, c))   # i comandi hanno la precedenza
                try:
                    job = IMGQ.get_nowait()
                except queue.Empty:
                    break
                if time.time() - job["t"] > 8.0:
                    job["event"].set()
                    continue
                try:
                    mac_export(order, job)
                except Exception as e:
                    job["result"] = None
                    job["error"] = "%s: %s" % (type(e).__name__, e)
                    job["event"].set()
            shows, order = mac_read_shows()
            if not shows:
                cache_clear()
        except Exception as e:
            log("PowerPoint Mac: %s" % e)
            fail_image_jobs()
            shows, order = [], []
            time.sleep(1.5)
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


START_PAGE = """<!DOCTYPE html><html lang="en" dir="ltr"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1"><title>__NAME__</title>
<link rel="icon" type="image/png" href="static/deckwhisperer-64.png">
<style>
:root{--bg:#f4f6f7;--card:#fff;--ink:#14262b;--mut:#5b6b70;--acc:#14434d;--ok:#1b7f4a;--bd:#d9e0e2}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font-family:Segoe UI,Arial,"Microsoft YaHei","PingFang SC","Noto Sans CJK SC",sans-serif}
.wrap{max-width:980px;margin:0 auto;padding:20px 16px 28px}
header{display:flex;align-items:center;gap:14px;margin-bottom:18px;flex-wrap:wrap}
header img{width:64px;height:64px}h1{font-size:26px;margin:0;letter-spacing:.02em}header small{color:var(--mut);display:block;margin-top:2px}
.guide{margin-inline-start:auto;font-size:16px;padding:10px 18px;text-decoration:none;background:var(--acc);color:#fff;border-radius:8px}
.grid{display:grid;grid-template-columns:minmax(260px,380px) 1fr;gap:18px}
@media(max-width:760px){.grid{grid-template-columns:1fr}}
.card{background:var(--card);border:1px solid var(--bd);border-radius:12px;padding:18px}
.card h2{font-size:17px;margin:0 0 12px}
.qr{text-align:center}.qr img{width:100%;max-width:320px;height:auto}
.url{font-size:20px;font-weight:700;word-break:break-all;margin:8px 0 2px;direction:ltr}.hint{color:var(--mut);font-size:14px}
.nic{margin-top:14px;text-align:start}.nic label{font-weight:700;font-size:14px;display:block;margin-bottom:5px}
.nic select{width:100%;font-size:15px;padding:8px;border:1px solid var(--bd);border-radius:8px;background:#fff}
.opt{display:flex;gap:12px;align-items:flex-start;padding:11px 4px;border-bottom:1px solid var(--bd);cursor:pointer}
.lockopt{background:#fff4f2;border-radius:8px;margin-top:6px;border-bottom:0}
.opt:last-of-type{border-bottom:0}.opt input{width:22px;height:22px;margin-top:2px;accent-color:var(--acc);flex:none}
.opt b{display:block;font-size:16px}.opt span{color:var(--mut);font-size:14px}
.pin{display:flex;gap:10px;align-items:center;margin-top:14px;flex-wrap:wrap}
.pin input{font-size:18px;padding:8px 10px;width:150px;border:1px solid var(--bd);border-radius:8px}
button{font-size:15px;padding:9px 14px;border:0;border-radius:8px;background:var(--acc);color:#fff;cursor:pointer}
#saved{color:var(--ok);font-size:14px;min-height:20px;margin-top:10px}
#shows{margin:0;padding-inline-start:18px;color:var(--ink)}#shows li{margin:2px 0}
.cr{font-size:13px;color:var(--mut);line-height:1.6;text-align:center;margin-top:22px}.cr a{color:var(--acc)}
.cr img{vertical-align:middle;margin-inline-end:4px}.links{margin-top:10px;font-size:14px}.links a{color:var(--acc)}
.remrow{display:flex;gap:18px;align-items:flex-start;flex-wrap:wrap}.remrow img{width:200px;height:200px;border:1px solid var(--bd);border-radius:8px;background:#fff}
button.stop{background:#8a3b2e}.warnp{color:#a4381f;font-weight:600}.okp{color:var(--ok)}
</style></head><body><div class="wrap">
<header><img src="static/deckwhisperer-192.png" alt="__NAME__"><div><h1>__NAME__</h1>
<small><span data-i18n="info_ver">Version</span> __VERSION__ &middot; Remote notes &amp; control for PowerPoint</small></div>
<a class="guide" id="guide" href="/guida" target="_blank">&#128214; <span data-i18n="g_guide">Guide</span></a></header>
<div class="grid">
<div class="card qr"><h2 data-i18n="g_step1">1. Scan the QR</h2><img id="qrimg" src="/qr.svg" alt="QR">
<div class="url" id="url">__URL__</div><div class="hint" data-i18n="g_hint">Phone connected to the same Wi-Fi network as the PC</div>
<div class="nic"><label for="nic" data-i18n="g_nic">Network adapter of the QR</label><select id="nic"></select>
<div class="hint" style="margin-top:5px" data-i18n="g_nic_hint">Pick the adapter your phones are connected to (e.g. the production network).</div></div>
<div class="nic"><label for="langsel" data-i18n="g_lang">Language</label><select id="langsel"></select>
<div class="hint" style="margin-top:5px" data-i18n="g_lang_hint">Default language of the phones (each phone can choose its own)</div></div></div>
<div class="card"><h2 data-i18n="g_step2">2. What the phones will show</h2>
<label class="opt"><input type="checkbox" id="u_bar"><div><b data-i18n="o_bar">Toolbar</b><span data-i18n="o_bar_d"></span></div></label>
<label class="opt"><input type="checkbox" id="u_style"><div><b data-i18n="o_style">Text style</b><span data-i18n="o_style_d"></span></div></label>
<label class="opt"><input type="checkbox" id="u_cur"><div><b data-i18n="o_cur">Current slide</b><span data-i18n="o_cur_d"></span></div></label>
<label class="opt"><input type="checkbox" id="u_next"><div><b data-i18n="o_next">Next slide</b><span data-i18n="o_next_d"></span></div></label>
<label class="opt"><input type="checkbox" id="u_goto"><div><b data-i18n="o_goto">Go to slide</b><span data-i18n="o_goto_d"></span></div></label>
<label class="opt"><input type="checkbox" id="u_ctl"><div><b data-i18n="o_ctl">Navigation</b><span id="ctl_d" data-i18n="o_ctl_d"></span></div></label>
<label class="opt" id="focusopt"><input type="checkbox" id="u_focus"><div><b data-i18n="o_focus">Keep PowerPoint in front</b><span data-i18n="o_focus_d"></span></div></label>
<label class="opt lockopt"><input type="checkbox" id="u_lock"><div><b data-i18n="o_lock">Lock remote screen</b><span data-i18n="o_lock_d"></span></div></label>
<div class="pin"><label for="pin"><b data-i18n="pin_label">PIN for the controls (optional)</b></label>
<input id="pin" maxlength="20" inputmode="numeric" autocomplete="off"><button id="savepin" data-i18n="pin_save">Save PIN</button></div>
<div id="saved"></div>
<h2 style="margin-top:14px" data-i18n="shows_title">Presentations in progress</h2><ul id="shows"><li data-i18n="shows_none"></li></ul>
<div class="links"><a href="/" target="_blank" data-i18n="open_notes">Open the notes page on this PC</a> &middot; <a href="/diagnostica" target="_blank" data-i18n="diag">Diagnostics</a></div>
</div></div>
<div class="card" id="remcard" style="margin-top:18px"><h2 data-i18n="r_title">3. Remote access (outside this network)</h2>
<div class="hint" data-i18n="r_desc"></div>
<div id="remoff" style="margin-top:12px"><button id="remstart" data-i18n="r_start">Switch on remote access</button></div>
<div id="remon" style="display:none;margin-top:12px"><div class="remrow"><img id="remqr" alt="QR">
<div><div class="url" id="remurl" style="font-size:16px"></div>
<div style="margin-top:10px"><button id="remcopy" data-i18n="r_copy">Copy link</button> <button id="remstop" class="stop" data-i18n="r_stop">Switch off</button></div>
<div id="remstat" class="hint" style="margin-top:10px"></div><div id="rempin" class="hint warnp"></div></div></div></div>
<div class="hint" style="margin-top:12px" data-i18n="r_privacy"></div></div>
<div class="cr">__CREDITS__</div></div>
<script>
const I18N = __I18N__;
const $ = id => document.getElementById(id), K = ["bar","style","cur","next","goto","ctl","lock"];
let lang = "en", noControl = false, autoIp = "";
const tr = k => ((I18N.langs[lang] || {})[k]) || I18N.langs.en[k] || k;
function applyLang(){
  document.documentElement.lang = lang; document.documentElement.dir = I18N.rtl.indexOf(lang) >= 0 ? "rtl" : "ltr";
  document.querySelectorAll("[data-i18n]").forEach(e => { e.textContent = tr(e.getAttribute("data-i18n")); });
  $("pin").placeholder = tr("pin_ph"); document.title = "__NAME__ - " + tr("g_title");
  $("guide").href = "/guida?lang=" + lang;
  if ($("credby")) $("credby").textContent = tr("credits_by");
  if (noControl) $("ctl_d").textContent = tr("no_control_note");
  const au = $("nic").options[0]; if (au) au.textContent = tr("g_nic_auto") + (autoIp ? " - " + autoIp : "");
  shows();
}
let timer = null;
function flash(k){ $("saved").textContent = tr(k); clearTimeout(timer); timer = setTimeout(() => { $("saved").textContent = ""; }, 2500); }
async function save(){
  const ui = {}; K.forEach(k => { ui[k] = $("u_" + k).checked; });
  try {
    const r = await fetch("/settings", { method:"POST", headers:{ "Content-Type":"application/json" },
      body: JSON.stringify({ ui: ui, pin: $("pin").value.trim(), nic: $("nic").value, lang: $("langsel").value, focus: $("u_focus").checked }) });
    flash(r.ok ? "saved" : "invalid");
    if (r.ok) { lang = $("langsel").value; applyLang(); load(false); }
  } catch(e) { flash("unreach"); }
}
function load(first){
  fetch("/settings.json").then(r => r.json()).then(d => {
    if (first) {
      K.forEach(k => { $("u_" + k).checked = k === "lock" ? d.ui[k] === true : d.ui[k] !== false; });
      $("pin").value = d.pin || "";
      $("u_focus").checked = d.focus !== false; $("focusopt").style.display = d.focus_ok === false ? "none" : "";
      const sel = $("nic"); sel.innerHTML = ""; autoIp = d.auto_ip || "";
      const auto = document.createElement("option"); auto.value = ""; sel.appendChild(auto);
      d.nics.forEach(n => { const o = document.createElement("option"); o.value = n.ip; o.textContent = n.name + " - " + n.ip; sel.appendChild(o); });
      sel.value = d.nic || ""; if (sel.value !== (d.nic || "")) sel.value = "";
      const ls = $("langsel"); ls.innerHTML = "";
      Object.keys(d.langs).forEach(c => { const o = document.createElement("option"); o.value = c; o.textContent = d.langs[c]; ls.appendChild(o); });
      ls.value = d.lang; lang = d.lang; noControl = !!d.no_control;
      $("u_ctl").disabled = noControl; $("u_goto").disabled = noControl;
      applyLang();
    }
    $("url").textContent = d.url; $("qrimg").src = "/qr.svg?t=" + Date.now();
  }).catch(() => flash("unreach"));
}
K.forEach(k => { $("u_" + k).onchange = save; });
$("u_focus").onchange = save;
$("nic").onchange = save; $("langsel").onchange = save;
$("savepin").onclick = save;
$("pin").addEventListener("keydown", e => { if (e.key === "Enter") save(); });
let lastShows = [];
function shows(){
  fetch("/state").then(r => r.json()).then(s => { lastShows = s.shows; drawShows(); }).catch(() => {});
}
function drawShows(){
  const ul = $("shows"); ul.innerHTML = "";
  if (!lastShows.length) { const li = document.createElement("li"); li.textContent = tr("shows_none"); ul.appendChild(li); return; }
  lastShows.forEach(x => { const li = document.createElement("li"); li.textContent = x.title + " - " + tr("slide") + " " + x.index + " / " + x.total; ul.appendChild(li); });
}
let remUrl = "";
async function remAct(a){ try { await fetch("/remote", { method:"POST", headers:{ "Content-Type":"application/json" }, body: JSON.stringify({ action: a }) }); } catch(e) {} remote(); }
function remote(){
  fetch("/remote.json").then(r => r.json()).then(d => {
    $("remoff").style.display = d.active ? "none" : "block"; $("remon").style.display = d.active ? "block" : "none";
    if (d.active) {
      $("remurl").textContent = d.url;
      if (remUrl !== d.url) { remUrl = d.url; $("remqr").src = "/remote-qr.svg?t=" + Date.now(); }
      $("remstat").textContent = tr("r_st_" + d.status) + (d.status === "error" && d.error ? " (" + d.error + ")" : "");
      $("rempin").className = "hint " + (d.pin_set ? "okp" : "warnp"); $("rempin").textContent = tr(d.pin_set ? "r_pin_ok" : "r_nopin");
    } else { remUrl = ""; }
  }).catch(() => {});
}
$("remstart").onclick = () => remAct("start"); $("remstop").onclick = () => remAct("stop");
$("remcopy").onclick = async () => { try { await navigator.clipboard.writeText(remUrl); } catch(e) {} flash("r_copied"); };
load(true); applyLang(); remote(); setInterval(remote, 2000);
shows(); setInterval(shows, 2000);
</script></body></html>"""


GUIDE_FILES = {c: "MANUAL-%s.md" % c for c in ("en", "it", "fr", "es", "ar", "zh")}


def guide_path(lang):
    name = GUIDE_FILES.get(lang, GUIDE_FILES["en"])
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
        mm = re.search(r"MANUAL-([a-z]{2})\.md$", url)
        if mm:
            url = "guida?lang=" + mm.group(1)
        if not url.startswith(("http://", "https://", "/", "guida")):
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


GUIDE_PAGE = """<!DOCTYPE html><html lang="__LANG__" dir="__DIR__"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1"><title>__NAME__ - __TITLE__</title>
<link rel="icon" type="image/png" href="static/deckwhisperer-64.png">
<style>body{margin:0;background:#f4f6f7;color:#14262b;font:16px/1.6 Segoe UI,Arial,"Microsoft YaHei","PingFang SC","Noto Sans CJK SC",sans-serif}
.wrap{max-width:860px;margin:0 auto;padding:18px 16px 40px;background:#fff;min-height:100vh}
.top{display:flex;gap:10px;align-items:center;flex-wrap:wrap;border-bottom:1px solid #d9e0e2;padding-bottom:10px;margin-bottom:6px}
.top img{width:44px;height:44px}.sw{margin-inline-start:auto;display:flex;gap:6px;flex-wrap:wrap}
.sw a{padding:4px 10px;border:1px solid #14434d;border-radius:14px;text-decoration:none;font-size:14px}
.sw a.cur{background:#14434d;color:#fff}
h1{font-size:26px}h2{font-size:21px;margin-top:28px}h3{font-size:17px}
table{border-collapse:collapse;width:100%;margin:10px 0}th,td{border:1px solid #d9e0e2;padding:7px 10px;text-align:start;vertical-align:top}
th{background:#eef3f4}code{background:#eef3f4;padding:1px 5px;border-radius:4px;direction:ltr;unicode-bidi:isolate}a{color:#14434d}</style></head><body><div class="wrap">
<div class="top"><img src="static/deckwhisperer-64.png" alt=""><b>__NAME__</b>
<span class="sw">__SWITCH__</span></div>
__BODY__</div></body></html>"""


def guide_page(lang):
    lang = lang if lang in GUIDE_FILES else "en"
    pth = guide_path(lang)
    if pth:
        with open(pth, encoding="utf-8") as f:
            body = md_to_html(f.read())
    else:
        body = "<p>Guide not found.</p>"
    names = I18N["names"]
    switch = "".join('<a href="guida?lang=%s"%s>%s</a>' % (c, ' class="cur"' if c == lang else "", html_escape(names.get(c, c)))
                     for c in GUIDE_FILES)
    title = I18N["langs"].get(lang, {}).get("g_guide", "Guide")
    return (GUIDE_PAGE.replace("__LANG__", lang).replace("__DIR__", "rtl" if lang in I18N["rtl"] else "ltr")
            .replace("__NAME__", html_escape(APP_NAME)).replace("__TITLE__", html_escape(title))
            .replace("__SWITCH__", switch).replace("__BODY__", body))


# -------------------------------------------------------------------- Server
# ---------------------------------------------------------------- accesso remoto (relay)
RELAY_URL = "https://remote.mefistofele.com:8443"   # server di rilancio (da cambiare con --relay)
REMOTE_GET_OK = (re.compile(r"^/$"), re.compile(r"^/guida$"), re.compile(r"^/favicon\.ico$"), re.compile(r"^/qr\.svg$"),
                 re.compile(r"^/static/[A-Za-z0-9_.-]{1,60}$"), re.compile(r"^/slide/[A-Za-z0-9_-]{1,32}/\d{1,5}\.png$"))
REMOTE_FAILS = deque(maxlen=20)       # tentativi di PIN sbagliati arrivati da fuori (istanti)
REMOTE_FAIL_MAX = 5                   # dopo 5 PIN sbagliati...
REMOTE_FAIL_WINDOW = 300              # ...in 5 minuti i comandi da fuori si bloccano finche' non passano


def remote_locked():
    """True se da fuori sono stati provati troppi PIN sbagliati da poco."""
    now = time.time()
    return sum(1 for t in REMOTE_FAILS if now - t < REMOTE_FAIL_WINDOW) >= REMOTE_FAIL_MAX


def remote_state():
    """Stato inviato ai telefoni remoti: con il PIN sempre richiesto per i comandi; senza PIN, solo lettura."""
    with COND:
        st = json.loads(json.dumps(STATE))
    st["pin_required"] = True
    st["remote"] = True
    if not PIN:
        st["control"] = False
        st["ui"]["ctl"] = False
        st["ui"]["goto"] = False
    return st


class RemoteLink:
    """Collegamento in uscita verso il relay: il PC chiede lavoro e il relay gli passa le richieste dei telefoni."""

    def __init__(self):
        self.lock = threading.Lock()
        self.active = False
        self.sid = ""
        self.key = ""
        self.status = "off"          # off | starting | online | error
        self.error = ""
        self.relay = RELAY_URL
        self.stop_evt = threading.Event()
        self.jobs_sem = threading.Semaphore(6)

    def url(self):
        return ("%s/s/%s/" % (self.relay.rstrip("/"), self.sid)) if self.active else ""

    def info(self):
        with self.lock:
            return {"active": self.active, "status": self.status, "error": self.error, "url": self.url(),
                    "pin_set": bool(PIN), "relay": self.relay}

    def start(self):
        import secrets
        with self.lock:
            if self.active:
                return
            self.sid = secrets.token_urlsafe(18)        # 24 caratteri: il "codice" nel link
            self.key = secrets.token_urlsafe(24)        # chiave segreta che resta solo sul PC
            self.active = True
            self.status = "starting"
            self.error = ""
            self.stop_evt = threading.Event()
            ev = self.stop_evt
        log("Accesso remoto attivato (%s)" % self.relay)
        threading.Thread(target=self._poll_loop, args=(ev, self.sid, self.key), daemon=True).start()
        threading.Thread(target=self._state_loop, args=(ev, self.sid, self.key), daemon=True).start()

    def stop(self):
        with self.lock:
            if not self.active:
                return
            self.active = False
            self.status = "off"
            self.error = ""
            ev, sid, key = self.stop_evt, self.sid, self.key
        ev.set()
        with COND:
            COND.notify_all()
        try:
            self._call("POST", "/pc/%s/bye" % sid, key, b"", timeout=5)
        except Exception:
            pass
        log("Accesso remoto disattivato")

    # --- rete
    def _call(self, method, path, key, body=None, ctype="application/octet-stream", timeout=40):
        import urllib.request
        req = urllib.request.Request(self.relay.rstrip("/") + path, data=body, method=method)
        req.add_header("X-DW-Key", key)
        req.add_header("User-Agent", "DeckWhisperer/" + APP_VERSION)
        if body is not None:
            req.add_header("Content-Type", ctype)
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return r.status, r.read(), r.headers

    def _set(self, status, error=""):
        with self.lock:
            if self.active:
                self.status, self.error = status, error

    def _poll_loop(self, ev, sid, key):
        delay = 1.0
        while not ev.is_set():
            try:
                code, data, _ = self._call("GET", "/pc/%s/poll" % sid, key, timeout=40)
                self._set("online")
                delay = 1.0
                if code == 200 and data:
                    job = json.loads(data.decode("utf-8"))
                    self.jobs_sem.acquire()
                    threading.Thread(target=self._serve, args=(sid, key, job), daemon=True).start()
            except Exception as e:
                if ev.is_set():
                    break
                code = getattr(e, "code", 0)
                self._set("error", ("HTTP %s" % code) if code else "%s" % type(e).__name__)
                log("Accesso remoto: %s (riprovo)" % (("HTTP %s" % code) if code else repr(e)[:120]))
                ev.wait(delay)
                delay = min(delay * 2, 15.0)

    def _state_loop(self, ev, sid, key):
        last = -1
        sent_at = 0.0
        while not ev.is_set():
            with COND:
                if VERSION == last and time.time() - sent_at < 20:
                    COND.wait(timeout=5)
                ver = VERSION
            if ev.is_set():
                break
            if ver == last and time.time() - sent_at < 20:
                continue
            try:
                self._call("POST", "/pc/%s/state" % sid, key, json.dumps(remote_state(), ensure_ascii=False).encode("utf-8"),
                           "application/json", timeout=15)
                last, sent_at = ver, time.time()
                self._set("online")
            except Exception as e:
                code = getattr(e, "code", 0)
                self._set("error", ("HTTP %s" % code) if code else "%s" % type(e).__name__)
                ev.wait(2.0)

    def _serve(self, sid, key, job):
        """Esegue sul server locale la richiesta di un telefono remoto e rimanda la risposta al relay."""
        try:
            status, ctype, body = self._local(job)
        except Exception as e:
            status, ctype, body = 502, "text/plain; charset=utf-8", ("Local error: %s" % type(e).__name__).encode("utf-8")
        finally:
            self.jobs_sem.release()
        try:
            self._call("POST", "/pc/%s/resp/%s?status=%d" % (sid, int(job["id"]), status), key, body, ctype, timeout=20)
        except Exception as e:
            log("Accesso remoto: risposta non consegnata (%s)" % type(e).__name__)

    def _local(self, job):
        import http.client
        method = str(job.get("method", ""))
        path = str(job.get("path", ""))
        route = path.split("?")[0]
        if method == "GET":
            ok = any(r.match(route) for r in REMOTE_GET_OK)
        else:
            ok = method == "POST" and route == "/cmd"
        if not ok or len(path) > 300 or any(c in path for c in "\r\n "):
            return 404, "text/plain; charset=utf-8", b"Not found"
        body = base64.b64decode(job["body"]) if job.get("body") else None
        if body is not None and len(body) > 4096:
            return 413, "text/plain; charset=utf-8", b"Too big"
        conn = http.client.HTTPConnection("127.0.0.1", PORT, timeout=15)
        try:
            hdr = {"X-DW-Remote": "1"}
            if body is not None:
                hdr["Content-Type"] = "application/json"
            conn.request(method, path, body=body, headers=hdr)
            r = conn.getresponse()
            data = r.read(6 * 1024 * 1024)
            return r.status, r.getheader("Content-Type", "application/octet-stream"), data
        finally:
            conn.close()


REMOTE = RemoteLink()


class Handler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"

    def log_message(self, *a):
        pass

    def _is_local(self):
        """True se la richiesta arriva dal PC stesso (le impostazioni si cambiano solo da qui)."""
        hd = getattr(self, "headers", None)
        if hd is not None and hd.get("X-DW-Remote"):      # arriva dal relay (telefono fuori rete): mai "locale"
            return False
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
                    job["error"] = "PowerPoint did not answer within 6 seconds"
                    log("Anteprima slide %s/%s: %s" % (sid, index, job["error"]))
        if not data:
            msg = (job["error"] if not DEMO_MODE else "") or "Preview not available"
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
                self._json(403, {"ok": False, "error": "local only"})
                return
            try:
                ok = apply_settings(json.loads(raw.decode("utf-8") or "{}"))
            except Exception:
                ok = False
            self._json(200 if ok else 400, {"ok": ok})
            return
        if path == "/remote":
            if not self._is_local():
                self._json(403, {"ok": False, "error": "local only"})
                return
            try:
                act = json.loads(raw.decode("utf-8") or "{}").get("action")
            except Exception:
                act = None
            if act == "start":
                REMOTE.start()
            elif act == "stop":
                REMOTE.stop()
            else:
                self._json(400, {"ok": False})
                return
            self._json(200, dict(REMOTE.info(), ok=True))
            return
        if path != "/cmd":
            self._json(404, {"ok": False, "error": "not found"})
            return
        if not CONTROL:
            self._json(403, {"ok": False, "error": "controls disabled"})
            return
        try:
            body = json.loads(raw.decode("utf-8") or "{}")
        except Exception:
            self._json(400, {"ok": False, "error": "invalid request"})
            return
        remote = bool(self.headers.get("X-DW-Remote"))
        if remote:                              # da fuori rete: PIN sempre obbligatorio e blocco dopo troppi errori
            if not PIN:
                self._json(403, {"ok": False, "error": "controls disabled"})
                return
            if remote_locked():
                self._json(429, {"ok": False, "error": "pin"})
                return
        if PIN and str(body.get("pin", "")) != PIN:
            if remote:
                REMOTE_FAILS.append(time.time())
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
                self._json(400, {"ok": False, "error": "invalid slide number"})
                return
        CMDQ.put(cmd)
        self._json(200, {"ok": True})

    def do_GET(self):
        path = self.path.split("?")[0]
        if path in ("/", "/index.html"):
            self._send(200, "text/html; charset=utf-8",
                       PAGE.replace("__I18N__", I18N_JSON).replace("__CREDITS__", CREDITS_HTML).replace("__VERSION__", APP_VERSION).encode("utf-8"))
        elif path.startswith("/slide/"):
            m = SLIDE_RE.match(path)
            if not m:
                log("Indirizzo anteprima non riconosciuto (404): %s" % self.path[:150])
                self._send(404, "text/plain; charset=utf-8", b"Not found")
                return
            self._slide(m.group(1), int(m.group(2)))
        elif path.startswith("/static/") or path == "/favicon.ico":
            name = "deckwhisperer.ico" if path == "/favicon.ico" else path[len("/static/"):]
            if name not in ASSET_FILES:
                self._send(404, "text/plain; charset=utf-8", b"Not found")
                return
            try:
                with open(resource_path("assets", name), "rb") as f:
                    data = f.read()
            except OSError:
                self._send(404, "text/plain; charset=utf-8", b"Not found")
                return
            self._send(200, ASSET_FILES[name], data)
        elif path == "/diagnostica":
            if "solo=info" in self.path:
                self._send(200, "text/plain; charset=utf-8", diagnostica_info().encode("utf-8"))
            else:
                self._send(200, "text/html; charset=utf-8", diagnostica_page().encode("utf-8"))
        elif path == "/qr.svg":
            if self.headers.get("X-DW-Remote"):  # richiesta dal relay: il QR e' quello del link remoto
                url = REMOTE.url()
                if not url:
                    self._send(404, "text/plain; charset=utf-8", b"Not found")
                    return
                self._send(200, "image/svg+xml", qr_svg(url).encode("utf-8"))
            else:
                self._send(200, "image/svg+xml", qr_svg(BASE_URL).encode("utf-8"))
        elif path in ("/remote.json", "/remote-qr.svg"):
            if not self._is_local():
                self._send(403, "text/plain; charset=utf-8", b"Local only")
                return
            if path == "/remote.json":
                self._json(200, REMOTE.info())
            elif REMOTE.url():
                self._send(200, "image/svg+xml", qr_svg(REMOTE.url()).encode("utf-8"))
            else:
                self._send(404, "text/plain; charset=utf-8", b"Not found")
        elif path in ("/start", "/qr"):
            if not self._is_local():
                self._send(403, "text/plain; charset=utf-8", "Management page: it only opens on the PC running DeckWhisperer".encode("utf-8"))
                return
            self._send(200, "text/html; charset=utf-8", START_PAGE.replace("__I18N__", I18N_JSON).replace("__URL__", BASE_URL).replace("__NAME__", html_escape(APP_NAME))
                       .replace("__VERSION__", APP_VERSION).replace("__CREDITS__", CREDITS_HTML).encode("utf-8"))
        elif path == "/guida":
            q = self.path.split("?", 1)[1] if "?" in self.path else ""
            mm = re.search(r"(?:^|&)lang=([a-z]{2})(?:&|$)", q)
            lang = mm.group(1) if mm and mm.group(1) in GUIDE_FILES else LANG
            self._send(200, "text/html; charset=utf-8", guide_page(lang).encode("utf-8"))
        elif path == "/settings.json":
            if not self._is_local():
                self._send(403, "text/plain; charset=utf-8", b"Local only")
                return
            self._send(200, "application/json; charset=utf-8",
                       json.dumps({"ui": UI, "pin": PIN, "no_control": NO_CONTROL, "url": BASE_URL,
                                   "nics": list_nics(), "nic": NIC_IP, "auto_ip": primary_ip(),
                                   "lang": LANG, "langs": I18N["names"],
                                   "focus": FOCUS_KEEP, "focus_ok": sys.platform == "win32"}).encode("utf-8"))
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
            self._send(404, "text/plain; charset=utf-8", b"Not found")


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
<html lang="en" dir="ltr">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>MEFISTOFELE DeckWhisperer</title>
<link rel="icon" type="image/png" sizes="64x64" href="static/deckwhisperer-64.png">
<link rel="icon" type="image/png" sizes="32x32" href="static/deckwhisperer-32.png">
<link rel="apple-touch-icon" href="static/apple-touch-icon.png">
<meta name="theme-color" content="#14434d">
<style>
  :root { --fg:#ffffff; --bg:#000000; }
  * { box-sizing: border-box; }
  html, body { margin:0; height:100%; background:var(--bg); color:var(--fg); }
  body { display:flex; flex-direction:column; font-family: Arial, sans-serif; overflow-y:auto; }
  #bar { flex:none; display:flex; flex-wrap:wrap; gap:8px 14px; align-items:center; padding:6px 10px;
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
  #main { flex:1 0 50vh; overflow:auto; padding:2vh 3vw; }   /* barra aperta: la pagina scende, le note restano almeno meta schermo */
  #notes { white-space:pre-wrap; overflow-wrap:anywhere; unicode-bidi:plaintext; text-align:start;
           font-family: Arial, "Microsoft YaHei", "PingFang SC", "Noto Sans CJK SC", "Noto Naskh Arabic", sans-serif; font-size:48px; line-height:1.3; letter-spacing:0; }
  #notes.wait { opacity:.5; font-style:italic; }
  #ctl { flex:none; position:sticky; bottom:0; z-index:6; display:flex; gap:10px; align-items:stretch; padding:8px 10px; background:#222;
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
  /* barra spenta dal PC: restano solo gli elementi scelti nella pagina di gestione, nessuna modifica possibile */
  body.nobar #bar > *:not(#showlbl):not(#info):not(.grp) { display:none !important; }
  body.nobar #strip, body.nobar #lockov { display:none !important; }
  body.nonav #prev, body.nonav #next { display:none; }
  /* Blocca schermo Remoto: nessun comando, solo lettura e scorrimento delle note */
  body.remlock #bar, body.remlock #strip, body.remlock #ctl, body.remlock #lockov,
  body.remlock #infoov, body.remlock #qrov { display:none !important; }
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
  body.hidebar #strip, body.locked #strip { display:flex; }
  #strip button { min-height:44px; padding:6px 8px; border-radius:22px; border:2px solid var(--c); background:transparent;
                  color:var(--c); font:600 14px Arial, sans-serif; cursor:pointer; white-space:nowrap; flex:1 1 auto; }
  #strip button.on { background:var(--c); color:#fff; }
  #sbar { --c:#c8a23a; } #sfull { --c:#12a8a8; } #sctl { --c:#1fae62; } #slock { --c:#ff5a5a; }
  body.nostripctl #sctl { display:none; }
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
  <button id="sbar" title="Reopen the toolbar">&#9776; <span id="t_sbar">Toolbar</span></button>
  <button id="sfull" title="Full screen">&#9974; <span id="t_sfull">Screen</span></button>
  <button id="sctl" class="on" title="Show/hide Navigation and Go to slide">&#9664;&#9654; <span id="t_sctl">Navigation</span></button>
  <button id="slock" title="Lock the screen">&#128274; <span id="t_slock">Lock</span></button>
  <button id="sunlock" title="Press and hold for 5 seconds to unlock"><span class="prog"></span><span>&#128275; <span id="t_sunlock">Press and hold for 5 seconds to unlock</span></span></button>
</div>
<div id="lockov"></div>
<div id="bar">
  <label id="showlbl"><span id="t_show">Presentation</span>
    <select id="show"></select>
  </label>
  <span class="grp">
  <label><span id="t_font">Font</span>
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
  <label><span id="t_size">Size</span>
    <input type="number" id="size" min="10" max="300" step="2" inputmode="numeric">
    <button id="sizeDown" title="Smaller">a&minus;</button>
    <button id="sizeUp" title="Bigger">A+</button>
  </label>
  <label><span id="t_text">Text</span> <input type="color" id="fg"></label>
  <label><span id="t_bg">Background</span> <input type="color" id="bg"></label>
  <label><span id="t_lh">Line spacing</span> <input type="range" id="lh" min="0.9" max="3" step="0.05"> <span id="lhv"></span></label>
  <label><span id="t_ls">Letter spacing</span> <input type="range" id="ls" min="-2" max="20" step="0.5"> <span id="lsv"></span></label>
  <button id="reset">Reset</button>
  </span>
  <div class="sec" id="sec1">What to show</div>
  <div class="chips">
    <button id="styleb" class="tg c-style" title="Show/hide the text style controls">Text style</button>
    <button id="ctlb" class="tg c-ctl" title="Show/hide Navigation and Go to slide">Navigation</button>
    <button id="curb" class="tg c-cur" title="Show/hide the current slide">Current slide</button>
    <button id="nextb" class="tg c-next" title="Show/hide the next slide">Next slide</button>
  </div>
  <div class="sec" id="sec2">Tools</div>
  <div class="chips">
    <button id="full" class="act c-full">Full screen</button>
    <button id="qrb" class="act c-qr">QR</button>
    <button id="infob" class="act c-info">Info</button>
    <button id="lockb" class="act c-lock">&#128274; <span id="t_lockb">Lock screen</span></button>
    <button id="hide" class="act c-hide">Hide toolbar</button>
  </div>
  <div id="info"><span id="pos"></span><span id="dot" title="Presentation"></span></div>
</div>
<div id="slides">
  <figure id="figcur"><figcaption id="cap_cur">On air</figcaption><img id="imgcur" alt=""><div class="ph" id="phcur"></div></figure>
  <figure id="fignext"><figcaption id="cap_next">Next</figcaption><img id="imgnext" alt=""><div class="ph" id="phnext"></div></figure>
</div>
<div id="main"><div id="notes" class="wait" dir="auto">Waiting for the presentation…</div></div>
<div id="ctl">
  <button id="prev" class="big">◀ <span id="t_prev">Previous</span></button>
  <div class="mid">
    <span id="ctlpos">–</span>
    <span class="goto"><input type="number" id="gonum" min="1" inputmode="numeric"> <button id="go">Go</button></span>
    <span id="ctlmsg"></span>
  </div>
  <button id="next" class="big"><span id="t_next">Next</span> ▶</button>
</div>
<div id="infoov">
  <div id="infobox">
    <img class="logo" src="static/deckwhisperer-192.png" alt="DeckWhisperer">
    <div class="ver"><span id="t_ver">Version</span> __VERSION__</div>
    <div>__CREDITS__</div>
    <div style="margin-top:8px;font-size:13px"><a id="t_guide" href="guida" target="_blank">Guide</a> &middot; <a id="t_diag" href="/diagnostica" target="_blank">Diagnostics</a></div>
    <div id="t_close" style="margin-top:12px;color:#888;font-size:12px">Tap to close</div>
  </div>
</div>
<div id="qrov" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.85);z-index:9;align-items:center;justify-content:center">
  <img src="qr.svg" alt="QR" style="width:min(80vw,80vh);height:min(80vw,80vh);background:#fff">
</div>

<script>
const DEF = { font:"Arial, sans-serif", size:48, fg:"#ffffff", bg:"#000000", lh:1.3, ls:0 };
let S = Object.assign({}, DEF);
try { Object.assign(S, JSON.parse(localStorage.getItem("pptNotesStyle") || "{}")); } catch(e) {}
const $ = id => document.getElementById(id);
const notes = $("notes");
const lsGet = (k, d) => { try { const v = localStorage.getItem(k); return v === null ? d : v; } catch(e) { return d; } };
const lsSet = (k, v) => { try { localStorage.setItem(k, v); } catch(e) {} };

/* ---- lingua: inglese di default; la sceglie il PC dalla pagina di gestione e vale per tutti i telefoni ---- */
const I18N = __I18N__;
let lang = "en";
const tr = k => ((I18N.langs[lang] || {})[k]) || I18N.langs.en[k] || k;
const TXT = { t_sbar:"s_bar", t_sfull:"s_full", t_sctl:"s_nav", t_slock:"s_lock", t_sunlock:"s_unlock", t_show:"present",
  t_font:"font", t_size:"size", t_text:"text", t_bg:"bg", t_lh:"lh", t_ls:"ls", reset:"reset", sec1:"sec_show", sec2:"sec_tools",
  styleb:"b_style", ctlb:"b_nav", curb:"b_cur", nextb:"b_next", full:"b_full", qrb:"b_qr", infob:"b_info", t_lockb:"b_lock", hide:"b_hide",
  cap_cur:"cap_cur", cap_next:"cap_next", t_prev:"c_prev", t_next:"c_next", go:"c_go", t_ver:"info_ver", t_guide:"info_guide",
  t_diag:"info_diag", t_close:"info_close", credby:"credits_by" };
const TTL = { sbar:"tt_s_bar", sfull:"tt_s_full", slock:"tt_s_lock", sunlock:"s_unlock", sizeDown:"tt_down", sizeUp:"tt_up",
  styleb:"tt_style", ctlb:"tt_nav", curb:"tt_cur", nextb:"tt_next", dot:"present" };
function pickLang() {
  // la lingua e' quella scelta nella pagina di gestione sul PC
  const l = (st && st.lang) || "en";
  return I18N.langs[l] ? l : "en";
}
function applyLang() {
  lang = pickLang();
  document.documentElement.lang = lang;
  document.documentElement.dir = I18N.rtl.indexOf(lang) >= 0 ? "rtl" : "ltr";
  Object.keys(TXT).forEach(id => { $(id).textContent = tr(TXT[id]); });
  Object.keys(TTL).forEach(id => { $(id).setAttribute("title", tr(TTL[id])); });
}

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
let st = { lang: "en", shows: [], pin_required: false, control: true, ui: { bar: true, style: true, cur: true, next: true, goto: true, ctl: true, lock: false } };
const uiOn = k => !st.ui || st.ui[k] !== false;
const remLock = () => !!(st.ui && st.ui.lock === true);   // "Blocca schermo Remoto": solo note, nessun comando
// vista fissa: barra spenta dal PC o schermo remoto bloccato -> conta solo la scelta della pagina di gestione
const fixedView = () => !uiOn("bar") || remLock();
const eff = (k, dev) => uiOn(k) && (fixedView() || dev);   // elemento visibile: scelta del PC + scelta del telefono   // scelto dal PC nella pagina di avvio
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
    notes.textContent = tr("wait");
    notes.classList.add("wait");
    $("pos").textContent = ""; $("ctlpos").textContent = "–";
    lastKey = null;
    updateSlides(null);
    return;
  }
  notes.classList.remove("wait");
  $("pos").textContent = cur.title + "  —  " + tr("slide") + " " + cur.index + " / " + cur.total;
  $("ctlpos").textContent = cur.index + " / " + cur.total;
  $("gonum").max = cur.total;
  notes.textContent = cur.notes && cur.notes.trim() ? cur.notes : tr("no_notes");
  const k = cur.id + "#" + cur.index;
  if (k !== lastKey) { $("main").scrollTop = 0; lastKey = k; }
  updateSlides(cur);
}

/* ---- anteprima slide in onda e successiva ---- */
function setSlide(img, ph, on, id, index, title, emptyText) {
  if (!on) return;
  if (!id || !index) { img.style.display = "none"; ph.textContent = emptyText; ph.style.display = "flex"; return; }
  const src = "slide/" + encodeURIComponent(id) + "/" + index + ".png?v=" + encodeURIComponent(title);
  if (img.getAttribute("data-src") !== src) { img.setAttribute("data-src", src); img.src = src; }
  ph.style.display = "none"; img.style.display = "block";
}
function updateSlides(cur) {
  const wait = tr("wait");
  setSlide($("imgcur"), $("phcur"), eff("cur", curOn), cur ? cur.id : "", cur ? cur.index : 0, cur ? cur.title : "", wait);
  const hasNext = !!cur && cur.index < cur.total;
  setSlide($("imgnext"), $("phnext"), eff("next", nextOn), hasNext ? cur.id : "", hasNext ? cur.index + 1 : 0,
           cur ? cur.title : "", cur ? tr("last") : wait);
}
["imgcur", "imgnext"].forEach(id => {
  $(id).onerror = () => {
    const ph = $(id === "imgcur" ? "phcur" : "phnext");
    const src = $(id).getAttribute("data-src");
    $(id).style.display = "none"; $(id).removeAttribute("data-src");
    ph.textContent = tr("prev_na"); ph.style.display = "flex";
    // chiede al programma il motivo e lo mostra al posto dell'anteprima
    if (src) fetch(src).then(r => r.text()).then(t => {
      if (t) ph.textContent = tr("prev_na") + ": " + t.slice(0, 220);
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
  if (!cur) { msg(tr("m_nopres")); return; }
  try {
    const r = await fetch("cmd", { method:"POST", headers:{ "Content-Type":"application/json" },
      body: JSON.stringify({ id: cur.id, action: action, n: n, pin: pin }) });
    if (r.status === 403) {
      let err = ""; try { err = (await r.json()).error; } catch(e) {}
      if (err === "pin") {
        if (retry) { msg(tr("m_pin_wrong")); pin = ""; lsSet("pptPin", ""); return; }
        const p = prompt(tr("m_pin_prompt"));
        if (p === null) return;
        pin = p; lsSet("pptPin", pin);
        return sendCmd(action, n, true);
      }
      msg(tr("m_disabled")); return;
    }
    if (!r.ok) msg(tr("m_err"));
  } catch(e) { msg(tr("m_lost")); }
}
$("next").onclick = () => sendCmd("next");
$("prev").onclick = () => sendCmd("prev");
$("go").onclick = () => { const n = parseInt($("gonum").value); if (n > 0) sendCmd("goto", n); };
$("gonum").addEventListener("keydown", e => { if (e.key === "Enter") $("go").click(); });
document.addEventListener("keydown", e => {
  if (!navOn() || e.target.tagName === "INPUT" || e.target.tagName === "SELECT") return;
  if (["ArrowRight","ArrowDown","PageDown"," "].includes(e.key)) { e.preventDefault(); sendCmd("next"); }
  else if (["ArrowLeft","ArrowUp","PageUp"].includes(e.key)) { e.preventDefault(); sendCmd("prev"); }
});
/* ---- cosa mostrare: ogni dispositivo ricorda le sue scelte ---- */
let styleOn = lsGet("pptStyle", "1") === "1";
let ctlOn = lsGet("pptCtl", "1") === "1";
let curOn = lsGet("pptCur", "1") === "1";
let nextOn = lsGet("pptNext", "1") === "1";
function navOn() { return st.control !== false && !remLock() && uiOn("ctl") && eff("ctl", ctlOn); }
function gotoOn() { return st.control !== false && uiOn("goto") && !remLock() && (fixedView() || ctlOn); }
function applyCtl() {
  const b = document.body.classList;
  const nav = navOn(), go = gotoOn(), allowed = st.control !== false && uiOn("ctl");
  b.toggle("remlock", remLock());
  b.toggle("nobar", !uiOn("bar"));
  if (!uiOn("bar")) b.remove("hidebar");
  b.toggle("nostyle", remLock() || !eff("style", styleOn));
  b.toggle("nonav", !nav);
  b.toggle("nogoto", !go);
  b.toggle("noctl", !nav && !go);
  b.toggle("nostripctl", !allowed && !(st.control !== false && uiOn("goto")));
  $("sctl").classList.toggle("on", ctlOn);
  b.toggle("showcur", eff("cur", curOn));
  b.toggle("shownext", eff("next", nextOn));
  const navAvail = st.control !== false && (uiOn("ctl") || uiOn("goto"));
  $("styleb").style.display = uiOn("style") ? "" : "none";
  $("ctlb").style.display = navAvail ? "" : "none";
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
  const es = new EventSource("events");
  es.onmessage = e => { try { st = JSON.parse(e.data); applyLang(); applyCtl(); render(); } catch(err) {} };
  es.onerror = () => { $("dot").className = ""; };
}
applyLang();
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
    ap.add_argument("--no-browser", action="store_true", help="do not open the management page on the PC")
    ap.add_argument("--relay", default="", help="relay address for remote access (default %s)" % RELAY_URL)
    ap.add_argument("--remote", action="store_true", help="switch remote access on at start")
    args = ap.parse_args()
    if args.relay:
        REMOTE.relay = args.relay.strip().rstrip("/")
    NO_CONTROL = bool(args.no_control)
    DEMO_MODE = bool(args.demo)
    saved_pin = load_settings()                  # scelte fatte l'ultima volta dalla pagina di avvio
    PIN = str(args.pin or saved_pin or "")
    CONTROL = ui_state()["ctl"]

    if not args.demo and IS_MAC:
        import shutil
        if not shutil.which("osascript"):
            print("ERROR: osascript not found (macOS with PowerPoint for Mac is required). Use --demo to try without PowerPoint.")
            return
    elif not args.demo:
        try:
            import win32com.client  # noqa: F401
        except Exception:
            print("ERROR: Windows with pywin32 and PowerPoint installed is required.")
            print("To try without PowerPoint use --demo.")
            input("Press Enter to close...")
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
        print("ERROR: no free port between %d and %d." % (args.port, args.port + 19))
        input("Premi Invio per chiudere...")
        return
    srv.daemon_threads = True

    PORT = port
    update_base_url()
    target = poll_demo if args.demo else (poll_powerpoint_mac if IS_MAC else poll_powerpoint)
    threading.Thread(target=target, args=(args.interval,), daemon=True).start()

    try:
        import ctypes
        ctypes.windll.kernel32.SetConsoleTitleW("%s %s" % (APP_NAME, APP_VERSION))
    except Exception:
        pass
    print("=" * 60)
    print(" %s %s - Remote notes & control for PowerPoint" % (APP_NAME, APP_VERSION))
    print("=" * 60)
    print(" Address for phones and other PCs on the network:")
    print("   " + BASE_URL)
    others = ["http://%s:%d/" % (i, port) for i in local_ips() if ("http://%s:%d/" % (i, port)) != BASE_URL]
    if others:
        print(" Other possible addresses (if there are several network cards):")
        for o in others:
            print("   " + o)
    print()
    try:
        print(qr_console(BASE_URL))
    except Exception as e:
        print(" (console QR not available: %s)" % e)
    print()
    print(" Management page (QR, network card, language and what to show on phones): http://localhost:%d/start" % port)
    if not CONTROL:
        print(" Next/previous controls: DISABLED")
    elif PIN:
        print(" Next/previous controls: enabled, protected by PIN")
    else:
        print(" Next/previous controls: enabled, NO PIN (anyone on the network can use them;")
        print("   set a PIN in the management page to protect them)")
    print(" On first run Windows asks to allow network access:")
    print(" choose 'Private networks' (and 'Public networks' if your Wi-Fi is seen as public).")
    print(" Error log: " + log_path())
    print(" Diagnostics: http://localhost:%d/diagnostica" % port)
    print(" Close this window to stop the program.")
    print()
    print(" MEFISTOFELE DeckWhisperer - Created and developed by www.andreabattaglioli.it")
    print(" Powered by www.mywebtvlive.com - Made in Bolognina")

    log("Avvio %s %s - indirizzo %s" % (APP_NAME, APP_VERSION, BASE_URL))
    if args.remote:
        REMOTE.start()
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
        input("Error. Press Enter to close...")
