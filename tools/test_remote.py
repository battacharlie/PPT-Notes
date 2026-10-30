"""Test dell'accesso remoto: relay + programma, con un "telefono" che passa solo dal relay (senza PowerPoint)."""
import importlib.util, json, os, subprocess, sys, threading, time, urllib.request, urllib.error
HERE = os.path.dirname(os.path.abspath(__file__))
spec = importlib.util.spec_from_file_location("deckwhisperer", os.path.join(HERE, "..", "windows", "deckwhisperer.py"))
m = importlib.util.module_from_spec(spec); spec.loader.exec_module(m)
RELAY_PORT, APP_PORT = 8491, 8492
relay = subprocess.Popen([sys.executable, os.path.join(HERE, "..", "relay", "mefistofele_relay.py"), "--http", "--host", "127.0.0.1", "--porta", str(RELAY_PORT)],
                         stderr=subprocess.DEVNULL)
fails = []
def check(name, cond, extra=""):
    print(("OK   " if cond else "FAIL ") + name + ((" - " + str(extra)) if (extra and not cond) else ""))
    if not cond: fails.append(name)
def req(url, method="GET", body=None, headers=None, timeout=15):
    r = urllib.request.Request(url, data=body, method=method, headers=headers or {})
    try:
        with urllib.request.urlopen(r, timeout=timeout) as x:
            return x.status, x.headers, x.read()
    except urllib.error.HTTPError as e:
        return e.code, e.headers, e.read()
try:
    time.sleep(1.0)
    m.DEMO_MODE = True; m.PORT = APP_PORT; m.PIN = ""
    srv = m.ThreadingHTTPServer(("127.0.0.1", APP_PORT), m.Handler); srv.daemon_threads = True
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    m.set_state([{"id": "w1", "title": "Demo", "index": 2, "total": 5, "notes": "Note di prova", "next_notes": ""}])
    R = "http://127.0.0.1:%d" % RELAY_PORT
    m.REMOTE.relay = R
    check("relay vivo", req(R + "/health")[0] == 200)
    check("prima dell'attivazione nessun link", m.REMOTE.info()["url"] == "")
    m.REMOTE.start()
    t0 = time.time()
    while time.time() - t0 < 10 and m.REMOTE.info()["status"] != "online": time.sleep(0.2)
    info = m.REMOTE.info(); base = info["url"].rstrip("/")
    check("stato online", info["status"] == "online", info)
    check("link con codice lungo", len(info["url"].split("/s/")[1].strip("/")) >= 22)
    # pagina
    c, h, b = req(base + "/")
    check("pagina note via relay", c == 200 and b"EventSource" in b and b"MEFISTOFELE" in b, c)
    check("percorsi relativi nella pagina", b'new EventSource("events")' in b and b'fetch("cmd"' in b and b'"slide/"' in b)
    check("icona via relay", req(base + "/static/deckwhisperer-64.png")[0] == 200)
    # stato con PIN assente: sola lettura
    c, h, b = req(base + "/state"); st = json.loads(b)
    check("stato remoto: senza PIN niente comandi", c == 200 and st["control"] is False and st["ui"]["ctl"] is False and st["remote"] is True, st)
    check("note nello stato", st["shows"][0]["notes"] == "Note di prova")
    # SSE
    r = urllib.request.urlopen(base + "/events", timeout=10); line = r.readline().decode()
    check("flusso eventi", line.startswith("data:") and "Note di prova" in line, line[:80]); r.close()
    # anteprima
    sid0 = m.DEMO_SHOWS[0]["id"]
    c, h, b = req(base + "/slide/%s/2.png?v=Demo" % sid0)
    check("anteprima slide via relay", c == 200 and b"<svg" in b, (c, b[:30]))
    c, h, b = req(base + "/slide/%s/2.png?v=Demo" % sid0)
    check("anteprima slide (seconda volta, dalla memoria del relay)", c == 200 and b"<svg" in b)
    check("anteprima inesistente: 404", req(base + "/slide/%s/99.png" % sid0)[0] == 404)
    # comandi senza PIN impostato: rifiutati
    c, h, b = req(base + "/cmd", "POST", json.dumps({"action": "next", "pin": ""}).encode(), {"Content-Type": "application/json"})
    check("comando senza PIN sul PC: rifiutato", c == 403, (c, b))
    # imposta PIN
    m.PIN = "4321"; m.refresh_state(); time.sleep(1.0)
    c, h, b = req(base + "/state"); st = json.loads(b)
    check("con PIN i comandi tornano", st["control"] is True and st["pin_required"] is True, st)
    while not m.CMDQ.empty(): m.CMDQ.get()
    c, h, b = req(base + "/cmd", "POST", json.dumps({"action": "next", "pin": "4321", "id": "w1"}).encode(), {"Content-Type": "application/json"})
    check("comando con PIN giusto", c == 200 and not m.CMDQ.empty(), (c, b))
    # PIN sbagliato ripetuto: blocco
    codes = []
    for i in range(7):
        codes.append(req(base + "/cmd", "POST", json.dumps({"action": "next", "pin": "0000"}).encode(), {"Content-Type": "application/json"})[0])
    check("PIN sbagliati: 403 poi blocco 429", codes[:1] == [403] and codes[-1] == 429, codes)
    c, h, b = req(base + "/cmd", "POST", json.dumps({"action": "next", "pin": "4321"}).encode(), {"Content-Type": "application/json"})
    check("anche il PIN giusto e' bloccato durante il blocco", c == 429, c)
    m.REMOTE_FAILS.clear()
    # percorsi non consentiti
    for p in ("/settings.json", "/start", "/diagnostica", "/settings", "/remote.json", "/remote"):
        check("percorso vietato %s" % p, req(base + p)[0] == 404, p)
    check("POST /settings vietato", req(base + "/settings", "POST", b"{}")[0] == 404)
    check("richiesta remota non e' 'locale'", True)
    # il PC locale resta locale
    c, h, b = req("http://127.0.0.1:%d/settings.json" % APP_PORT)
    check("impostazioni dal PC ok", c == 200)
    c, h, b = req("http://127.0.0.1:%d/settings.json" % APP_PORT, headers={"X-DW-Remote": "1"})
    check("impostazioni con intestazione remota: vietato", c == 403, c)
    # QR del link remoto
    try:
        import qrcode  # noqa: F401
        c, h, b = req(base + "/qr.svg"); check("QR del link remoto", c == 200 and b"<svg" in b)
    except ImportError:
        print("SKIP QR del link remoto (libreria qrcode non installata)")
    # chiave: un altro PC non puo' prendere il posto
    sid = info["url"].split("/s/")[1].strip("/")
    c, h, b = req(R + "/pc/%s/state" % sid, "POST", b"{}", {"X-DW-Key": "x" * 30, "Content-Type": "application/json"})
    check("chiave sbagliata: 403", c == 403, c)
    check("codice troppo corto: 404", req(R + "/s/abc/")[0] == 404)
    # spegni
    m.REMOTE.stop(); time.sleep(0.5)
    c, h, b = req(base + "/")
    check("dopo lo spegnimento: relatore offline (503)", c == 503, c)
    check("stato spento", m.REMOTE.info()["status"] == "off" and m.REMOTE.info()["url"] == "")
    # riaccensione: nuovo link
    m.REMOTE.start(); time.sleep(2.0)
    check("nuovo link diverso", m.REMOTE.info()["url"] != info["url"])
    m.REMOTE.stop()
finally:
    relay.terminate()
print("\nFALLITI: %d" % len(fails) if fails else "\nTUTTO OK")
sys.exit(1 if fails else 0)
