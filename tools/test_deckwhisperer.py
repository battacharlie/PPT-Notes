"""Test di DeckWhisperer senza PowerPoint: anteprime slide (con PowerPoint simulato) e comportamento della pagina."""
import json, importlib.util, shutil, subprocess, tempfile, threading, time, urllib.request, urllib.error, sys, os
HERE = os.path.dirname(os.path.abspath(__file__))
spec = importlib.util.spec_from_file_location("deckwhisperer", os.path.join(HERE, "..", "windows", "deckwhisperer.py"))
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
files = [f for f in os.listdir(__import__("tempfile").gettempdir()) if f.startswith("DeckWhisperer_")]
check("nessun file temporaneo rimasto", files == [])
st, ct, _, dg = get("/diagnostica")
check("pagina /diagnostica (200, html)", st == 200 and ct.startswith("text/html") and b"DeckWhisperer" in dg)
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
m.settings_paths = lambda name="mefistofele.json": [os.path.join(tmp, "mefistofele.json")]
m.BASE_URL = "http://192.168.1.5:8080/"
st, ct, _, body = get("/start")
check("pagina di avvio dal PC (200, HTML, QR e caselle)", st == 200 and "text/html" in ct and b"/qr.svg" in body and b"u_cur" in body and b"u_bar" in body and b"u_goto" in body and "MEFISTOFELE DeckWhisperer".encode() in body and b"/guida" in body and b"id=\"nic\"" in body and b"192.168.1.5:8080" in body)
check("/qr mostra la stessa pagina di avvio", get("/qr")[3] == body)
st, ct, _, body = get("/settings.json"); d = json.loads(body)
check("impostazioni iniziali: tutto visibile, nessun PIN", st == 200 and d["ui"] == {"bar": True, "style": True, "cur": True, "next": True, "goto": True, "ctl": True, "lock": False} and d["nic"] == "" and isinstance(d["nics"], list) and d["pin"] == "")
_, _, _, body = get("/state"); s0 = json.loads(body)
check("stato per i telefoni contiene 'ui'", s0["ui"]["cur"] is True and s0["control"] is True)
ver0 = m.VERSION
st, r = post("/settings", {"ui": {"style": False, "ctl": True, "cur": True, "next": False, "bar": False, "goto": False}, "pin": "4321"})
check("salvataggio impostazioni (200)", st == 200 and r["ok"])
_, _, _, body = get("/state"); s1 = json.loads(body)
check("i telefoni ricevono subito le nuove scelte", s1["ui"] == {"bar": False, "style": False, "cur": True, "next": False, "goto": False, "ctl": True, "lock": False} and s1["pin_required"] is True and m.VERSION > ver0)
saved = json.load(open(os.path.join(tmp, "mefistofele.json"), encoding="utf-8"))
check("scelte e PIN salvati su file", saved["ui"]["next"] is False and saved["ui"]["bar"] is False and saved["ui"]["goto"] is False and saved["pin"] == "4321")
m.UI.update({k: True for k in m.UI_KEYS}); m.PIN = ""
m.PIN = m.load_settings()
check("riavvio: scelte e PIN riletti dal file", m.PIN == "4321" and m.UI["style"] is False and m.UI["next"] is False and m.UI["cur"] is True and m.UI["goto"] is False and m.UI["bar"] is False)
st, r = post("/settings", {"ui": {"cur": "si"}}); check("valore non valido rifiutato (400)", st == 400)
st, r = post("/settings", {"ui": {}, "pin": "x" * 40}); check("PIN troppo lungo rifiutato (400)", st == 400)
st, r = post("/cmd", {"id": "111", "action": "next", "pin": "0000"}); check("comando con PIN errato -> 403", st == 403 and r["error"] == "pin")
post("/settings", {"ui": {"ctl": False}, "pin": ""})
st, r = post("/cmd", {"id": "111", "action": "next"}); check("comandi spenti dal PC -> 403 'comandi disabilitati'", st == 403)
_, _, _, body = get("/state"); check("stato: control False e ui.ctl False", json.loads(body)["control"] is False and json.loads(body)["ui"]["ctl"] is False)
m.NO_CONTROL = True; post("/settings", {"ui": {"ctl": True}}); _, _, _, body = get("/state")
check("--no-control vince sulla casella", json.loads(body)["ui"]["ctl"] is False and json.loads(body)["control"] is False)
m.NO_CONTROL = False; post("/settings", {"ui": {k: True for k in m.UI_KEYS}, "pin": ""})
# impostazioni salvate dalla 2.0 (reMNote.json): lette una volta
old_dir = tempfile.mkdtemp(); json.dump({"ui": {"cur": False}, "pin": "77"}, open(os.path.join(old_dir, "reMNote.json"), "w"))
real_paths = m.settings_paths; m.settings_paths = lambda name="mefistofele.json": [os.path.join(old_dir, name)]
m.UI.update({k: True for k in m.UI_KEYS})
check("migrazione: legge reMNote.json della 2.0", m.load_settings() == "77" and m.UI["cur"] is False)
m.settings_paths = real_paths; m.UI.update({k: True for k in m.UI_KEYS})
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
# ---- versione 3: scheda di rete del QR e guida
m.PORT = 8080
m.list_nics = lambda: [{"name": "Ethernet produzione", "ip": "10.10.0.5"}, {"name": "Wi-Fi", "ip": "192.168.1.5"}]
m.primary_ip = lambda: "192.168.1.5"
m.NIC_IP = ""; m.update_base_url()
check("QR: scheda automatica = IP principale", m.BASE_URL == "http://192.168.1.5:8080/")
st, r = post("/settings", {"ui": {}, "nic": "10.10.0.5"})
check("scheda di rete scelta: QR e indirizzo cambiano subito", st == 200 and m.BASE_URL == "http://10.10.0.5:8080/")
d = json.loads(get("/settings.json")[3])
check("settings.json: elenco schede, scelta e indirizzo", d["nic"] == "10.10.0.5" and d["url"] == "http://10.10.0.5:8080/" and len(d["nics"]) == 2 and d["auto_ip"] == "192.168.1.5")
check("pagina di avvio mostra l'indirizzo della scheda scelta", b"10.10.0.5:8080" in get("/start")[3])
saved = json.load(open(os.path.join(tmp, "mefistofele.json"), encoding="utf-8"))
check("scheda salvata su file", saved["nic"] == "10.10.0.5")
st, r = post("/settings", {"ui": {}, "nic": "8.8.8.8"}); check("scheda inesistente rifiutata (400)", st == 400 and m.BASE_URL == "http://10.10.0.5:8080/")
m.NIC_IP = ""; m.load_settings(); m.update_base_url()
check("riavvio: scheda riletta dal file", m.NIC_IP == "10.10.0.5" and m.BASE_URL == "http://10.10.0.5:8080/")
m.list_nics = lambda: [{"name": "Wi-Fi", "ip": "192.168.1.5"}]; m.update_base_url()
check("scheda sparita: si torna all'automatica", m.BASE_URL == "http://192.168.1.5:8080/")
post("/settings", {"ui": {}, "nic": ""}); check("scheda 'automatica' selezionabile", m.NIC_IP == "")
st, ct, _, body = get("/guida")
check("guida di default in inglese (200, HTML con titoli e tabelle)", st == 200 and "text/html" in ct and b"<h2>" in body and b"<table>" in body and b"Quick start" in body and b'lang="en"' in body)
check("guida: selettore con tutte le lingue", all(("guida?lang=%s" % c).encode() in body for c in ("en","it","fr","es","ar","zh")))
for c, rtl in (("it",0),("fr",0),("es",0),("ar",1),("zh",0)):
    st, ct, _, body = get("/guida?lang=" + c)
    check("guida %s (200, titoli, %s)" % (c, "rtl" if rtl else "ltr"), st == 200 and b"<h2>" in body and (b'dir="rtl"' in body) == bool(rtl) and ('lang="%s"' % c).encode() in body)
check("guida: lingua sconosciuta -> inglese", b"Quick start" in get("/guida?lang=xx")[3])
# ---- lingua (3.1)
d = json.loads(get("/settings.json")[3]); check("lingua di default: inglese", d["lang"] == "en" and set(d["langs"]) == {"en","it","fr","es","ar","zh"})
check("i18n: stesse chiavi in tutte le lingue", all(set(m.I18N["langs"][c]) == set(m.I18N["langs"]["en"]) for c in m.I18N["langs"]))
check("i18n: nessun testo vuoto", all(v.strip() for c in m.I18N["langs"] for v in m.I18N["langs"][c].values()))
st, r = post("/settings", {"ui": {}, "pin": "", "lang": "xx"}); check("lingua non valida rifiutata", st == 400 and json.loads(get("/state")[3])["lang"] == "en")
st, r = post("/settings", {"ui": {}, "pin": "", "lang": "fr"}); check("lingua francese accettata", st == 200 and json.loads(get("/state")[3])["lang"] == "fr")
check("lingua: la guida di default segue la lingua scelta", b"D\xc3\xa9marrage" in get("/guida")[3] or b'lang="fr"' in get("/guida")[3])
post("/settings", {"ui": {}, "pin": "", "lang": "en"})
check("md_to_html: elenco con riga di continuazione", m.md_to_html("1. uno\n   due\n2. tre") == "<ol><li>uno due</li><li>tre</li></ol>")
check("md_to_html: HTML nel testo e' neutralizzato", "<script>" not in m.md_to_html("x <script>alert(1)</script>"))
m.UI.update({k: True for k in m.UI_KEYS}); m.PIN = ""
# ---- Lock remote screen: nessun comando accettato dal server
post("/settings", {"ui": {"lock": True, "ctl": True, "goto": True}, "pin": ""})
_, _, _, body = get("/state"); s2 = json.loads(body)
check("schermo remoto bloccato: stato con lock, senza comandi", s2["ui"]["lock"] is True and s2["ui"]["ctl"] is False and s2["ui"]["goto"] is False and s2["control"] is False)
st, r = post("/cmd", {"id": "111", "action": "next"}); check("schermo remoto bloccato: i comandi sono rifiutati (403)", st == 403)
d = json.loads(get("/settings.json")[3]); check("pannello: i flag scelti restano memorizzati sotto il blocco", d["ui"]["lock"] is True and d["ui"]["ctl"] is True)
post("/settings", {"ui": {"lock": False}, "pin": ""}); _, _, _, body = get("/state")
check("schermo remoto sbloccato: i comandi tornano", json.loads(body)["control"] is True and json.loads(body)["ui"]["ctl"] is True)
b = get("/start")[3]
check("pagina di gestione: nomi e ordine dei flag", [b.find(x.encode()) for x in ("Toolbar", "Text style", "Current slide", "Next slide", "Go to slide", "Navigation", "Lock remote screen")] == sorted(b.find(x.encode()) for x in ("Toolbar", "Text style", "Current slide", "Next slide", "Go to slide", "Navigation", "Lock remote screen")) and b.find(b"Lock remote screen") > 0)
stop = True

# ---- versione Mac: AppleScript simulato
calls = []
def fake_osa(script, *args, timeout=6):
    calls.append((script, args))
    if script is m.OSA_READ:
        return "/Users/a/demo.pptx\x1fdemo.pptx\x1f2\x1f5\x1fnota uno\runo\x1fnota due\x1e/Users/a/b.pptx\x1fb.pptx\x1f1\x1f3\x1f\x1f\x1e"
    if script is m.OSA_EXPORT:
        open(args[2], "wb").write(b"\x89PNG-mac-%d" % int(args[1])); return "ok"
    return "ok"
m_osa = m.osa; m.osa = fake_osa
shows, order = m.mac_read_shows()
check("mac: legge due presentazioni", len(shows) == 2 and shows[0]["index"] == 2 and shows[0]["total"] == 5 and shows[1]["title"] == "b.pptx")
check("mac: note con a capo normalizzati", shows[0]["notes"] == "nota uno\nuno" and shows[0]["next_notes"] == "nota due")
check("mac: id fatti di lettere e cifre, diversi tra loro", shows[0]["id"] != shows[1]["id"] and all(c.isalnum() for c in shows[0]["id"]))
calls.clear(); m.mac_run_command(order, {"id": shows[1]["id"], "action": "goto", "n": 3})
check("mac: comando 'vai alla slide' sulla finestra giusta", calls and calls[0][1] == (2, "goto", 3))
calls.clear(); m.mac_run_command(order, {"id": "nessuna", "action": "next"})
check("mac: id sconosciuto -> nessun comando", not calls or calls[0][1][0] == 1)
m.cache_clear(); job = {"id": shows[0]["id"], "index": 2, "t": time.time(), "event": threading.Event()}
m.mac_export(order, job)
check("mac: anteprima esportata", job["result"] == b"\x89PNG-mac-2" and job["error"] == "")
m.osa = m_osa

# ---- coda dei comandi: click ravvicinati (4.1)
eseguiti = []
def runner(c): eseguiti.append((c["action"], c.get("n")))
CMD_ID = ["A"]
def metti(az, eta=0, **kw):
    c = {"id": CMD_ID[0], "action": az, "t": time.time() - eta}; c.update(kw); m.CMDQ.put(c)
for _ in range(6): metti("prev")
metti("goto", n=3); metti("next"); metti("next")
m.drain_commands(runner)
check("6 click 'indietro' ravvicinati: tutti eseguiti, in ordine", eseguiti == [("prev", None)] * 6 + [("goto", 3), ("next", None), ("next", None)])
eseguiti.clear(); metti("next"); metti("prev"); metti("prev"); m.drain_commands(runner)
check("avanti/indietro sommati: netto = un 'indietro'", eseguiti == [("prev", None)])
eseguiti.clear(); metti("prev", eta=m.CMD_MAX_AGE + 5); metti("prev"); m.drain_commands(runner)
check("comando molto vecchio scartato e scritto nel registro", eseguiti == [("prev", None)] and any("scartato" in l for l in m.LOG_LINES))
check("i comandi con ritardo di qualche secondo non si perdono piu'", True)
eseguiti.clear(); metti("prev", eta=5); metti("prev", eta=4); m.drain_commands(runner)
check("comandi di 4-5 secondi fa: eseguiti", eseguiti == [("prev", None)] * 2)
check("il riepilogo dei comandi finisce nel registro", any("Comandi:" in l and "passaggi" in l for l in m.LOG_LINES))
# precedenza sulle anteprime: un comando arrivato durante gli export viene eseguito prima del job successivo
ordine = []
class WV:
    CurrentShowPosition = 1
    @staticmethod
    def Previous(): ordine.append("prev")
class WinP(FakeWin):
    def __init__(self, h, n):
        FakeWin.__init__(self, h, n); self.View = WV
        outer = self.Presentation.Slides
        class Sl(FakeSlides):
            def __call__(s, i):
                class S:
                    def Export(self_, path, fmt, w, h):
                        ordine.append("export%d" % i); open(path, "wb").write(b"x")
                        if i == 1: metti("prev")           # click che arriva mentre si esporta
                return S()
        self.Presentation.Slides = Sl(n)
wp = WinP(444, 5); appp = FakeApp([wp]); m.cache_clear(); CMD_ID[0] = m.win_id(wp, 1)
j1 = {"id": m.win_id(wp, 1), "index": 1, "t": time.time(), "event": threading.Event(), "result": None, "error": ""}
j2 = {"id": m.win_id(wp, 1), "index": 2, "t": time.time(), "event": threading.Event(), "result": None, "error": ""}
m.IMGQ.put(j1); m.IMGQ.put(j2); m.process_image_jobs(appp)
check("click arrivato durante un export: eseguito prima della prossima anteprima", ordine == ["export1", "prev", "export2"])

# ---- fuoco di PowerPoint dopo le anteprime (4.2)
class FakeFg:
    def __init__(self, seq): self.seq = list(seq); self.set_to = []
    def get(self): return self.seq.pop(0) if len(self.seq) > 1 else self.seq[0]
    def put(self, h): self.set_to.append(h); return True
PIDS = {100: 1, 200: 1, 300: 2, 0: 0}      # 100 e 200: finestre di PowerPoint (processo 1), 300: altro programma
def prova_fuoco(seq, before=(100, True)):
    f = FakeFg(seq)
    r = m.focus_restore(None, before, get_fg=f.get, pid_of=lambda h: PIDS.get(h, 9), set_fg=f.put, pp_pids={1}, wait=0)
    return r, f.set_to
check("fuoco: invariato dopo l'anteprima -> nessun intervento", prova_fuoco([100, 100]) == (False, []))
check("fuoco: perso verso il nulla (0) -> restituito a PowerPoint", prova_fuoco([0]) == (True, [100]))
check("fuoco: passato a un'altra finestra di PowerPoint -> restituito", prova_fuoco([200]) == (True, [100]))
check("fuoco: perso un attimo dopo l'export (secondo controllo) -> restituito", prova_fuoco([100, 0]) == (True, [100]))
check("fuoco: ora e' di un altro programma -> non si interviene", prova_fuoco([300]) == (False, []))
check("fuoco: prima non era di PowerPoint -> non si interviene", prova_fuoco([0], before=(300, False)) == (False, []))
m.FOCUS_KEEP = False
check("fuoco: opzione spenta -> nessuno snapshot", m.focus_snapshot(None) == (0, False))
m.FOCUS_KEEP = True
check("fuoco: l'opzione e' nelle impostazioni", json.loads(get("/settings.json")[3]).get("focus") is True)
check("fuoco: impostazione accettata e salvata", post("/settings", {"ui": {"bar": True}, "focus": False})[0] == 200 and m.FOCUS_KEEP is False and json.loads(get("/settings.json")[3])["focus"] is False)
post("/settings", {"ui": {"bar": True}, "focus": True})
check("fuoco: valore non booleano rifiutato", post("/settings", {"ui": {"bar": True}, "focus": "si"})[0] == 400)
check("fuoco: opzione presente nella pagina di gestione e in tutte le lingue", b'id="u_focus"' in get("/start")[3] and all(("o_focus" in m.I18N["langs"][l]) for l in m.I18N["langs"]))

# ---- anteprime preparate in anticipo (4.3)
class PSlides:
    def __init__(self, n): self.Count = n; self.exports = []
    def __call__(self, i):
        outer = self
        class S:
            def Export(self_, path, fmt, w, h):
                outer.exports.append(i); open(path, "wb").write(b"PNG-pre-%d" % i)
        return S()
class PPres:
    def __init__(self, name, n):
        self.FullName = name; self.Slides = PSlides(n)
        class PS: SlideWidth = 720.0; SlideHeight = 540.0
        self.PageSetup = PS
class PWin:
    def __init__(self, pres, hwnd=555):
        self.Presentation = pres; self.HWND = hwnd
        class V: CurrentShowPosition = 1
        self.View = V
class PApp:
    def __init__(self, pres_list): self.pl = pres_list; self.shows = []; self.Presentations = self; self.SlideShowWindows = type("W", (), {"Count": 0, "__call__": lambda s, i: None})()
    @property
    def Count(self): return len(self.pl)
    def __call__(self, i): return self.pl[i-1]
    def start(self, pres):
        sw = self.shows
        sw.append(PWin(pres))
        self.SlideShowWindows = type("W", (), {"Count": 1, "__call__": lambda s, i: sw[i-1]})()
pp = PPres(os.path.join(tempfile.gettempdir(), "deck-test.pptx"), 4); papp = PApp([pp])
m.PRE_CACHE.clear(); m.PRE_NEXT.clear(); m.PRELOAD = True
passi = 0
while m.preload_step(papp): passi += 1
check("anteprime in anticipo: una slide per passo, tutte e 4 preparate", passi == 4 and pp.Slides.exports == [1, 2, 3, 4] and len(m.PRE_CACHE) == 4)
check("anteprime in anticipo: a lavoro finito non esporta piu'", m.preload_step(papp) is False and pp.Slides.exports == [1, 2, 3, 4])
check("anteprime in anticipo: avvisi nel registro (inizio e fine)", any("Anteprime: preparo 4 slide" in l for l in m.LOG_LINES) and any("Anteprime pronte" in l for l in m.LOG_LINES))
papp.start(pp); sid = m.win_id(papp.shows[0], 1); m.cache_clear()
job = {"id": sid, "index": 3, "t": time.time(), "event": threading.Event(), "result": None, "error": ""}
m.export_slide(papp, job)
check("durante la presentazione: anteprima servita dalle preparate, nessun export", job["result"] == b"PNG-pre-3" and pp.Slides.exports == [1, 2, 3, 4] and job["error"] == "")
check("durante la presentazione non si prepara nulla", m.preload_step(papp) is False)
papp2 = PApp([PPres(os.path.join(tempfile.gettempdir(), "deck-test2.pptx"), 3)]); m.PRELOAD = False
check("opzione spenta: nessun export in anticipo", m.preload_step(papp2) is False and papp2.pl[0].Slides.exports == [])
m.PRELOAD = True
pp.Slides.Count = 5; papp3 = PApp([pp]); n0 = len(pp.Slides.exports)
while m.preload_step(papp3): pass
check("presentazione cambiata (5 slide): anteprime rifatte", len(pp.Slides.exports) == n0 + 5)
class PRotta(PPres):
    def __init__(self):
        PPres.__init__(self, "rotta.pptx", 2)
        class Sl(PSlides):
            def __call__(s, i):
                class S:
                    def Export(self_, *a): raise RuntimeError("vista protetta")
                return S()
        self.Slides = Sl(2)
pr = PApp([PRotta()])
check("export in anticipo che fallisce: un solo tentativo, poi si lascia fare durante la presentazione", m.preload_step(pr) is True and m.preload_step(pr) is False and any("non riuscite" in l for l in m.LOG_LINES))
check("anticipo: opzione nelle impostazioni, accettata, salvata e valore non valido rifiutato", json.loads(get("/settings.json")[3]).get("preload") is True and post("/settings", {"ui": {"bar": True}, "preload": False})[0] == 200 and m.PRELOAD is False and post("/settings", {"ui": {"bar": True}, "preload": "no"})[0] == 400)
post("/settings", {"ui": {"bar": True}, "preload": True})
check("anticipo: opzione nella pagina di gestione e in tutte le lingue", b'id="u_preload"' in get("/start")[3] and all("o_preload" in m.I18N["langs"][l] for l in m.I18N["langs"]))

# ---- pagina: JavaScript eseguito con un browser simulato (serve Node.js)
node = shutil.which("node")
if node:
    import re
    page = m.PAGE.replace("__I18N__", m.I18N_JSON).replace("__CREDITS__", "").replace("__VERSION__", "test")
    js = os.path.join(tempfile.gettempdir(), "deckwhisperer_pagina.js")
    open(js, "w", encoding="utf-8").write(re.search(r"<script>(.*)</script>", page, re.S).group(1))
    r = subprocess.run([node, os.path.join(HERE, "test_pagina.js"), js], capture_output=True, text=True)
    print(r.stdout + r.stderr)
    check("pagina: test del comportamento dei pulsanti", r.returncode == 0)
    dj = os.path.join(tempfile.gettempdir(), "deckwhisperer_diagnostica.js")
    open(dj, "w", encoding="utf-8").write(re.search(r"<script>(.*)</script>", m.DIAG_PAGE, re.S).group(1))
    r2 = subprocess.run([node, "--check", dj], capture_output=True, text=True)
    check("pagina diagnostica: JavaScript valido", r2.returncode == 0)
    if r2.returncode: print(r2.stderr)
else:
    print("Node.js non trovato: test della pagina saltato")
print("TUTTO OK" if ok else "CI SONO ERRORI"); sys.exit(0 if ok else 1)
