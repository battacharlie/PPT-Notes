"""Test di reMNote senza PowerPoint: anteprime slide (con PowerPoint simulato) e comportamento della pagina."""
import json, importlib.util, shutil, subprocess, tempfile, threading, time, urllib.request, urllib.error, sys, os
HERE = os.path.dirname(os.path.abspath(__file__))
spec = importlib.util.spec_from_file_location("remnote", os.path.join(HERE, "..", "windows", "remnote.py"))
m = importlib.util.module_from_spec(spec); spec.loader.exec_module(m)

class FakeSlides:
    def __init__(self, n, fail_on=None): self.Count = n; self.exports = []; self.fail_on = fail_on
    def __call__(self, i):
        outer = self
        class S:
            def Export(self_, path, fmt, w, h):
                outer.exports.append((i, fmt, w, h))
                if outer.fail_on == i: raise RuntimeError("export fallito")
                open(path, "wb").write(b"\x89PNG-fake-slide-%d-%dx%d" % (i, w, h))
        return S()
class FakePres:
    def __init__(self, n, fail_on=None):
        self.Slides = FakeSlides(n, fail_on)
        class PS: SlideWidth = 720.0; SlideHeight = 540.0   # 4:3
        self.PageSetup = PS
class FakeWin:
    def __init__(self, hwnd, n, fail_on=None): self.HWND = hwnd; self.Presentation = FakePres(n, fail_on)
class FakeApp:
    def __init__(self, wins): self._w = wins; self.SlideShowWindows = self
    @property
    def Count(self): return len(self._w)
    def __call__(self, i): return self._w[i-1]

class FailAllSlides(FakeSlides):
    def __call__(self, i):
        class S:
            def Export(self_, path, fmt, w, h): raise RuntimeError("Slide.Export non consentito")
        return S()
class FakeWinView(FakeWin):
    """Export dalla presentazione fallisce; dalla finestra di presentazione (slide in onda = 1) riesce."""
    def __init__(self, hwnd, n):
        FakeWin.__init__(self, hwnd, n); self.Presentation.Slides = FailAllSlides(n)
        class V:
            CurrentShowPosition = 1
            class Slide:
                @staticmethod
                def Export(path, fmt, w, h): open(path, "wb").write(b"\x89PNG-da-finestra")
        self.View = V
w1 = FakeWin(111, 5); w2 = FakeWin(222, 3, fail_on=2); w3 = FakeWinView(333, 4)
app = FakeApp([w1, w2, w3])
m.DEMO_MODE = False
srv = m.ThreadingHTTPServer(("127.0.0.1", 8087), m.Handler); srv.daemon_threads = True
threading.Thread(target=srv.serve_forever, daemon=True).start()
stop = False
def poller():
    while not stop:
        m.process_image_jobs(app); time.sleep(0.05)
threading.Thread(target=poller, daemon=True).start()

def get(path):
    try:
        r = urllib.request.urlopen("http://127.0.0.1:8087" + path, timeout=10)
        return r.status, r.headers.get("Content-Type"), r.headers.get("Cache-Control"), r.read()
    except urllib.error.HTTPError as e:
        return e.code, e.headers.get("Content-Type"), None, e.read()

ok = True
def check(name, cond):
    global ok; ok &= bool(cond); print(("OK   " if cond else "FAIL ") + name)

t = time.time(); st, ct, cc, body = get("/slide/111/2.png")
check("slide 2 esportata (200, image/png)", st == 200 and ct == "image/png")
check("contenuto = file scritto da Export, 960x720 per un 4:3", body == b"\x89PNG-fake-slide-2-960x720")
check("formato PNG richiesto a PowerPoint", w1.Presentation.Slides.exports == [(2, "PNG", 960, 720)])
check("cache-control per il browser", cc == "private, max-age=300")
st, _, _, body2 = get("/slide/111/2.png")
check("seconda richiesta dalla cache (nessun nuovo Export)", st == 200 and body2 == body and len(w1.Presentation.Slides.exports) == 1)
st, *_ = get("/slide/111/9.png");   check("slide fuori intervallo -> 404", st == 404)
st, *_ = get("/slide/999/1.png");   check("presentazione sconosciuta -> 404", st == 404)
t = time.time(); st, _, _, err = get("/slide/222/2.png")
check("Export che fallisce -> 404 senza bloccare (%.1fs)" % (time.time()-t), st == 404 and time.time()-t < 3)
check("il motivo dell'errore arriva alla pagina", b"export fallito" in err and b"esportazione fallita" in err)
st, _, _, b9 = get("/slide/333/1.png"); check("secondo tentativo dalla finestra di presentazione (slide in onda)", st == 200 and b9 == b"\x89PNG-da-finestra")
st, _, _, e9 = get("/slide/333/2.png"); check("slide non in onda: errore spiegato", st == 404 and b"Slide.Export non consentito" in e9)
st, _, _, e8 = get("/slide/111/9.png"); check("fuori intervallo: messaggio chiaro", st == 404 and b"fuori intervallo (1-5)" in e8)
st, _, _, b3 = get("/slide/222/1.png"); check("altra presentazione funziona dopo l'errore", st == 200 and b"slide-1" in b3)
files = [f for f in os.listdir(__import__("tempfile").gettempdir()) if f.startswith("reMNote_")]
check("nessun file temporaneo rimasto", files == [])
st, ct, _, dg = get("/diagnostica")
check("pagina /diagnostica (200, html)", st == 200 and ct.startswith("text/html") and b"reMNote" in dg)
check("diagnostica: versione, percorso del registro e messaggi recenti", m.APP_VERSION.encode() in dg and b"Registro previsto" in dg and b"Anteprima slide" in dg)
st, ct, _, di = get("/diagnostica?solo=info")
check("diagnostica: testo semplice aggiornabile", st == 200 and ct.startswith("text/plain") and b"Ultimi messaggi" in di)
st, *_ = get("/pagina-che-non-esiste")
check("indirizzo sconosciuto -> 404 e registrato", st == 404 and any("pagina-che-non-esiste" in l for l in m.LOG_LINES))
st, *_ = get("/slide/%20/1.png")
check("indirizzo anteprima non valido -> 404 e registrato", st == 404 and any("anteprima non riconosciuto" in l for l in m.LOG_LINES))
check("registro scritto su file (accanto al programma o in temp)", m.LOG_FILE_USED[0] != "" and os.path.exists(m.LOG_FILE_USED[0]))
m.cache_clear(); check("cache svuotabile", m.cache_get((111, 2)) is None and len(m.SLIDE_CACHE) == 0)
for i in range(60): m.cache_put(("x", i), b"a")
check("cache limitata a %d voci" % m.SLIDE_CACHE_MAX, len(m.SLIDE_CACHE) == m.SLIDE_CACHE_MAX)
# ---- id finestra: HWND che arriva come metodo (caso reale visto su un PC) o assente
class HwndMetodo:
    def __init__(self, v, nome): self._v = v; self.Presentation = type("P", (), {"FullName": nome})()
    def HWND(self): return self._v
class HwndAssente:
    def __init__(self, nome): self.Presentation = type("P", (), {"FullName": nome})()
a, b = HwndMetodo(1182454, "C:\\x\\1.pptx"), HwndMetodo(99871, "C:\\x\\2.pptx")
check("id con HWND come metodo = numero", m.win_id(a, 1) == "1182454" and m.win_id(b, 2) == "99871")
c, d = HwndAssente("C:\\x\\1.pptx"), HwndAssente("C:\\x\\2.pptx")
check("id senza HWND: valido e diverso per presentazione", m.win_id(c, 1) != m.win_id(d, 2) and all(m.SLIDE_RE.match("/slide/%s/1.png" % m.win_id(w, i)) for i, w in ((1, c), (2, d))))
check("id sempre accettato dall'indirizzo anteprima", all(m.SLIDE_RE.match("/slide/%s/1.png" % m.win_id(w, 1)) for w in (a, b, c, d, object())))
# ---- pagina di avvio e impostazioni (versione 2)
import json
def post(path, obj):
    req = urllib.request.Request("http://127.0.0.1:8087" + path, data=json.dumps(obj).encode(), headers={"Content-Type": "application/json"})
    try:
        r = urllib.request.urlopen(req, timeout=10); return r.status, json.loads(r.read())
    except urllib.error.HTTPError as e:
        return e.code, json.loads(e.read())
tmp = tempfile.mkdtemp()
m.settings_paths = lambda: [os.path.join(tmp, "reMNote.json")]
m.BASE_URL = "http://192.168.1.5:8080/"
st, ct, _, body = get("/start")
check("pagina di avvio dal PC (200, HTML, QR e caselle)", st == 200 and "text/html" in ct and b"/qr.svg" in body and b"u_cur" in body and b"192.168.1.5:8080" in body)
check("/qr mostra la stessa pagina di avvio", get("/qr")[3] == body)
st, ct, _, body = get("/settings.json"); d = json.loads(body)
check("impostazioni iniziali: tutto visibile, nessun PIN", st == 200 and d["ui"] == {"style": True, "ctl": True, "cur": True, "next": True} and d["pin"] == "")
_, _, _, body = get("/state"); s0 = json.loads(body)
check("stato per i telefoni contiene 'ui'", s0["ui"]["cur"] is True and s0["control"] is True)
ver0 = m.VERSION
st, r = post("/settings", {"ui": {"style": False, "ctl": True, "cur": True, "next": False}, "pin": "4321"})
check("salvataggio impostazioni (200)", st == 200 and r["ok"])
_, _, _, body = get("/state"); s1 = json.loads(body)
check("i telefoni ricevono subito le nuove scelte", s1["ui"] == {"style": False, "ctl": True, "cur": True, "next": False} and s1["pin_required"] is True and m.VERSION > ver0)
saved = json.load(open(os.path.join(tmp, "reMNote.json"), encoding="utf-8"))
check("scelte e PIN salvati su file", saved["ui"]["next"] is False and saved["pin"] == "4321")
m.UI.update({k: True for k in m.UI_KEYS}); m.PIN = ""
m.PIN = m.load_settings()
check("riavvio: scelte e PIN riletti dal file", m.PIN == "4321" and m.UI["style"] is False and m.UI["next"] is False and m.UI["cur"] is True)
st, r = post("/settings", {"ui": {"cur": "si"}}); check("valore non valido rifiutato (400)", st == 400)
st, r = post("/settings", {"ui": {}, "pin": "x" * 40}); check("PIN troppo lungo rifiutato (400)", st == 400)
st, r = post("/cmd", {"id": "111", "action": "next", "pin": "0000"}); check("comando con PIN errato -> 403", st == 403 and r["error"] == "pin")
post("/settings", {"ui": {"ctl": False}, "pin": ""})
st, r = post("/cmd", {"id": "111", "action": "next"}); check("comandi spenti dal PC -> 403 'comandi disabilitati'", st == 403)
_, _, _, body = get("/state"); check("stato: control False e ui.ctl False", json.loads(body)["control"] is False and json.loads(body)["ui"]["ctl"] is False)
m.NO_CONTROL = True; post("/settings", {"ui": {"ctl": True}}); _, _, _, body = get("/state")
check("--no-control vince sulla casella", json.loads(body)["ui"]["ctl"] is False and json.loads(body)["control"] is False)
m.NO_CONTROL = False; post("/settings", {"ui": {k: True for k in m.UI_KEYS}, "pin": ""})
# una richiesta da un altro dispositivo non puo' cambiare le impostazioni
m.local_ips = lambda: []
class FakeSelf:
    client_address = ("192.168.1.77", 5555)
    _is_local = m.Handler._is_local
check("richiesta da un telefono non e' 'locale'", FakeSelf()._is_local() is False)
FakeSelf.client_address = ("127.0.0.1", 5555); check("richiesta dal PC e' 'locale'", FakeSelf()._is_local() is True)
FakeSelf.client_address = ("::1", 5555); check("IPv6 locale e' 'locale'", FakeSelf()._is_local() is True)
m.local_ips = lambda: ["192.168.1.5"]; FakeSelf.client_address = ("192.168.1.5", 5555)
check("indirizzo di rete del PC stesso e' 'locale'", FakeSelf()._is_local() is True)
stop = True

# ---- pagina: JavaScript eseguito con un browser simulato (serve Node.js)
node = shutil.which("node")
if node:
    import re
    page = m.PAGE.replace("__CREDITS__", "").replace("__VERSION__", "test")
    js = os.path.join(tempfile.gettempdir(), "remnote_pagina.js")
    open(js, "w", encoding="utf-8").write(re.search(r"<script>(.*)</script>", page, re.S).group(1))
    r = subprocess.run([node, os.path.join(HERE, "test_pagina.js"), js], capture_output=True, text=True)
    print(r.stdout + r.stderr)
    check("pagina: test del comportamento dei pulsanti", r.returncode == 0)
    dj = os.path.join(tempfile.gettempdir(), "remnote_diagnostica.js")
    open(dj, "w", encoding="utf-8").write(re.search(r"<script>(.*)</script>", m.DIAG_PAGE, re.S).group(1))
    r2 = subprocess.run([node, "--check", dj], capture_output=True, text=True)
    check("pagina diagnostica: JavaScript valido", r2.returncode == 0)
    if r2.returncode: print(r2.stderr)
else:
    print("Node.js non trovato: test della pagina saltato")
print("TUTTO OK" if ok else "CI SONO ERRORI"); sys.exit(0 if ok else 1)
