"""Test di reMNote senza PowerPoint: anteprime slide (con PowerPoint simulato) e comportamento della pagina."""
import importlib.util, shutil, subprocess, tempfile, threading, time, urllib.request, urllib.error, sys, os
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

w1 = FakeWin(111, 5); w2 = FakeWin(222, 3, fail_on=2)
app = FakeApp([w1, w2])
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
t = time.time(); st, *_ = get("/slide/222/2.png")
check("Export che fallisce -> 404 senza bloccare (%.1fs)" % (time.time()-t), st == 404 and time.time()-t < 3)
st, _, _, b3 = get("/slide/222/1.png"); check("altra presentazione funziona dopo l'errore", st == 200 and b"slide-1" in b3)
files = [f for f in os.listdir(__import__("tempfile").gettempdir()) if f.startswith("reMNote_")]
check("nessun file temporaneo rimasto", files == [])
m.cache_clear(); check("cache svuotabile", m.cache_get((111, 2)) is None and len(m.SLIDE_CACHE) == 0)
for i in range(60): m.cache_put(("x", i), b"a")
check("cache limitata a %d voci" % m.SLIDE_CACHE_MAX, len(m.SLIDE_CACHE) == m.SLIDE_CACHE_MAX)
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
else:
    print("Node.js non trovato: test della pagina saltato")
print("TUTTO OK" if ok else "CI SONO ERRORI"); sys.exit(0 if ok else 1)
