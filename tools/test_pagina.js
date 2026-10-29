const fs = require("fs"), vm = require("vm");
const code = fs.readFileSync(process.argv[2], "utf8");
function boot(seed) {
const els = {}; const store = Object.assign({}, seed); let es = null; const docL = {}; let fsCalls = 0; const vids = [];
function El(id) {
  const e = { id, style: { setProperty() {} }, attrs: {}, _cls: new Set(), textContent: "", value: "", options: [], className: "", max: 0 };
  e.classList = { add: c => e._cls.add(c), remove: c => e._cls.delete(c), contains: c => e._cls.has(c),
    toggle: (c, f) => { if (f === undefined) f = !e._cls.has(c); f ? e._cls.add(c) : e._cls.delete(c); return f; } };
  e.setAttribute = (k, v) => { e.attrs[k] = v; }; e.getAttribute = k => (k in e.attrs ? e.attrs[k] : null);
  e.removeAttribute = k => { delete e.attrs[k]; }; e.addEventListener = () => {};
  e.appendChild = c => { e.options.push(c); };
  Object.defineProperty(e, "innerHTML", { set(v) { e.options.length = 0; }, get() { return ""; } });
  e.scrollTop = 0; return e;
}
const $ = id => els[id] || (els[id] = El(id));
const body = El("body"), docEl = El("html"); docEl.style.setProperty = () => {};
const sandbox = {
  document: { getElementById: $, body, documentElement: Object.assign(docEl, { requestFullscreen: () => { fsCalls++; return Promise.resolve(); } }),
    createElement: tag => { const e = El(tag);
      if (tag === "video") { e.played = 0; e.paused = true; e.canPlayType = () => "maybe"; e.play = () => { e.played++; e.paused = false; return Promise.resolve(); }; vids.push(e); }
      return e; }, addEventListener(t, f) { (docL[t] = docL[t] || []).push(f); }, fullscreenElement: null, visibilityState: "visible" },
  navigator: {}, localStorage: { getItem: k => (k in store ? store[k] : null), setItem: (k, v) => { store[k] = String(v); } },
  EventSource: function () { es = this; }, fetch: (u) => Promise.resolve({ ok: true, status: 200, text: () => Promise.resolve(globalThis.__errText || "") }),
  prompt: () => null, setTimeout, clearTimeout, console, encodeURIComponent, parseInt, parseFloat, Number, JSON, Math,
};
vm.createContext(sandbox); vm.runInContext(code, sandbox);
return { $, body, store, sandbox, run: c => vm.runInContext(c, sandbox), vids: () => vids, tap: () => (docL.pointerdown || []).forEach(f => f({})), fs: () => fsCalls, send: st => es.onmessage({ data: JSON.stringify(st) }) };
}
let { $, body, store, sandbox, send } = boot({ pptCur: "0", pptNext: "0" });
let fails = 0; const T = (n, c) => { if (!c) fails++; console.log((c ? "OK   " : "FAIL ") + n); };
const sh = (idx, total = 4, id = "d1") => ({ id, title: "demo.pptx", index: idx, total, notes: "n" + idx, next_notes: "" });
const has = c => body._cls.has(c);

// stato iniziale: stile e comandi visibili, anteprime nascoste
send({ control: true, pin_required: false, shows: [sh(1)] });
T("iniziale: stile visibile", !has("nostyle")); T("iniziale: comandi visibili", !has("noctl"));
T("iniziale: anteprime nascoste", !has("showcur") && !has("shownext"));
T("iniziale: nessuna immagine caricata", !$("imgcur").src && !$("imgnext").src);
T("pulsanti Stile/Comandi evidenziati, anteprime no", $("styleb")._cls.has("on") && $("ctlb")._cls.has("on") && !$("curb")._cls.has("on") && !$("nextb")._cls.has("on"));

// slide in onda
$("curb").onclick();
T("slide in onda: attiva", has("showcur") && !has("shownext") && $("curb")._cls.has("on"));
T("slide in onda: carica /slide/d1/1.png", $("imgcur").src === "/slide/d1/1.png?v=demo.pptx");
T("slide in onda: la successiva non viene caricata", !$("imgnext").src);
// successiva
$("nextb").onclick();
T("successiva: carica /slide/d1/2.png", has("shownext") && $("imgnext").src === "/slide/d1/2.png?v=demo.pptx");
// avanza
send({ control: true, shows: [sh(2)] });
T("avanzando: in onda = 2, successiva = 3", $("imgcur").src.includes("/d1/2.png") && $("imgnext").src.includes("/d1/3.png"));
// ultima slide
send({ control: true, shows: [sh(4)] });
T("ultima slide: segnaposto 'Ultima slide' al posto dell'immagine", $("phnext").textContent === "Ultima slide" && $("phnext").style.display === "flex" && $("imgnext").style.display === "none");
// nessuna presentazione
send({ control: true, shows: [] });
T("senza presentazione: segnaposto attesa", $("phcur").textContent.includes("In attesa") && $("imgcur").style.display === "none");
send({ control: true, shows: [sh(1, 4, "d2")] });
T("nuova presentazione: immagine ricaricata", $("imgcur").src.includes("/slide/d2/1.png") && $("imgcur").style.display === "block");
// errore immagine
$("imgcur").onerror();
T("errore immagine: messaggio, niente immagine rotta", $("phcur").textContent === "Anteprima non disponibile" && $("imgcur").style.display === "none");

// motivo dell'errore mostrato al posto dell'anteprima
globalThis.__errText = "esportazione fallita (prova)";
sandbox.globalThis = globalThis;
$("imgcur").onerror();
setTimeout(() => {
  T("errore: il motivo compare al posto dell'anteprima", $("phcur").textContent === "Anteprima non disponibile: esportazione fallita (prova)");
  rest();
}, 20);
function rest() {
// nascondere
$("curb").onclick(); $("nextb").onclick();
T("anteprime di nuovo nascoste", !has("showcur") && !has("shownext"));
$("styleb").onclick();
T("stile testo nascosto", has("nostyle") && !$("styleb")._cls.has("on"));
$("ctlb").onclick();
T("comandi avanzamento nascosti", has("noctl") && !$("ctlb")._cls.has("on"));
$("styleb").onclick(); $("ctlb").onclick();
T("si possono rimostrare", !has("nostyle") && !has("noctl"));

// memoria per dispositivo
T("scelte salvate nel dispositivo", store.pptStyle === "1" && store.pptCtl === "1" && store.pptCur === "0" && store.pptNext === "0");
$("curb").onclick(); $("styleb").onclick();
T("scelte aggiornate nel dispositivo", store.pptCur === "1" && store.pptStyle === "0");

// comandi disabilitati dal server
send({ control: false, shows: [sh(1)] });
T("--no-control: area comandi nascosta e pulsante Comandi sparito", has("noctl") && $("ctlb").style.display === "none");

// ---- versione 2: cosa mostrare e' deciso dal PC (pagina di avvio)
{
  const D = boot({});   // dispositivo nuovo, senza scelte salvate
  const hasD = c => D.body._cls.has(c);
  D.send({ control: true, ui: { style: true, ctl: true, cur: true, next: true }, shows: [sh(1)] });
  T("v2: dispositivo nuovo, tutto abilitato dal PC -> tutto visibile", !hasD("nostyle") && !hasD("noctl") && hasD("showcur") && hasD("shownext"));
  T("v2: le anteprime si caricano da sole", D.$("imgcur").src.includes("/slide/d1/1.png") && D.$("imgnext").src.includes("/slide/d1/2.png"));
  D.send({ control: true, ui: { style: false, ctl: true, cur: false, next: true }, shows: [sh(1)] });
  T("v2: stile spento dal PC -> nascosto e pulsante sparito", hasD("nostyle") && D.$("styleb").style.display === "none");
  T("v2: slide in onda spenta dal PC -> nascosta e pulsante sparito", !hasD("showcur") && D.$("curb").style.display === "none");
  T("v2: le altre restano disponibili", hasD("shownext") && D.$("nextb").style.display === "" && !hasD("noctl") && D.$("ctlb").style.display === "");
  D.send({ control: true, ui: { style: true, ctl: false, goto: false, cur: true, next: false }, shows: [sh(2)] });
  T("v2: comandi spenti dal PC -> area e pulsante nascosti", hasD("noctl") && D.$("ctlb").style.display === "none");
  T("v2: slide successiva spenta dal PC -> nascosta", !hasD("shownext") && D.$("nextb").style.display === "none");
  D.send({ control: true, ui: { style: true, ctl: true, cur: true, next: true }, shows: [sh(2)] });
  T("v2: il PC riaccende tutto, subito", !hasD("nostyle") && !hasD("noctl") && hasD("showcur") && hasD("shownext"));
  D.$("curb").onclick();
  T("v2: sul telefono si puo' comunque nascondere ci\u00f2 che il PC permette", !hasD("showcur"));
  // ---- versione 3: barra, "vai alla slide", striscia, lucchetto
  const E = boot({});
  const hasE = c => E.body._cls.has(c);
  const U = (o) => Object.assign({ bar: true, style: true, cur: true, next: true, goto: true, ctl: true }, o);
  E.send({ control: true, ui: U({}), shows: [sh(1)] });
  T("v3: tutto abilitato -> barra, stile e 'vai' visibili", !hasE("nobar") && !hasE("nogoto") && !hasE("nostyle"));
  E.send({ control: true, ui: U({ goto: false }), shows: [sh(1)] });
  T("v3: 'Vai direttamente alla slide' spento dal PC -> nascosto, Avanti/Indietro restano", hasE("nogoto") && !hasE("noctl"));
  E.send({ control: true, ui: U({ bar: false }), shows: [sh(1)] });
  T("v3: barra spenta dal PC -> classe nobar", hasE("nobar"));
  E.send({ control: true, ui: U({}), shows: [sh(1)] });
  T("v3: barra riaccesa subito", !hasE("nobar") && !hasE("nogoto"));
  E.$("hide").onclick();
  T("v3: 'Nascondi barra' -> classe hidebar (striscia visibile)", hasE("hidebar"));
  E.$("sbar").onclick();
  T("v3: pulsante Barra della striscia riapre la barra", !hasE("hidebar"));
  E.send({ control: true, ui: U({ ctl: false, goto: false }), shows: [sh(1)] });
  T("v3: navigazione e vai spenti dal PC -> pulsante nella striscia nascosto", hasE("nostripctl"));
  E.send({ control: true, ui: U({}), shows: [sh(1)] });
  E.$("sctl").onclick();
  T("v3: striscia: pulsante Comandi nasconde Avanti/Indietro", hasE("noctl") && !E.$("sctl")._cls.has("on"));
  E.$("sctl").onclick();
  T("v3: e li rimostra", !hasE("noctl") && E.$("sctl")._cls.has("on"));
  // ---- Barra degli strumenti spenta: vista fissa decisa dalla pagina di gestione
  {
    const F = boot({ pptCur: "0", pptNext: "0", pptStyle: "0", pptCtl: "0" });   // il telefono aveva nascosto tutto
    const hasF = c => F.body._cls.has(c);
    F.send({ control: true, ui: U({ bar: false }), shows: [sh(2)] });
    T("fissa: gli elementi scelti compaiono anche se il telefono li aveva nascosti", hasF("showcur") && hasF("shownext") && !hasF("nostyle") && !hasF("noctl") && !hasF("nogoto") && !hasF("nonav"));
    T("fissa: le anteprime si caricano", F.$("imgcur").src.includes("/slide/d1/2.png") && F.$("imgnext").src.includes("/slide/d1/3.png"));
    F.$("hide").onclick();
    T("fissa: 'Nascondi barra' non ha effetto (nessuna modifica possibile)", !hasF("hidebar"));
    F.send({ control: true, ui: U({ bar: false, style: false, next: false, goto: false }), shows: [sh(2)] });
    T("fissa: gli elementi non scelti restano nascosti", hasF("nostyle") && !hasF("shownext") && hasF("nogoto") && hasF("showcur") && !hasF("nonav"));
    F.send({ control: true, ui: U({ bar: false, ctl: false }), shows: [sh(2)] });
    T("fissa: solo 'Vai alla Slide' -> area comandi visibile, senza Avanti/Indietro", hasF("nonav") && !hasF("noctl") && !hasF("nogoto"));
    F.send({ control: true, ui: U({ bar: false, ctl: false, goto: false }), shows: [sh(2)] });
    T("fissa: né Navigazione né Vai -> area comandi nascosta", hasF("noctl"));
    // Blocca schermo Remoto
    F.send({ control: true, ui: U({ lock: true, ctl: false, goto: false }), shows: [sh(2)] });
    T("schermo remoto bloccato: classe remlock, niente stile né comandi", hasF("remlock") && hasF("nostyle") && hasF("noctl"));
    T("schermo remoto bloccato: restano le anteprime scelte", hasF("showcur") && hasF("shownext"));
    F.send({ control: true, ui: U({ lock: true, cur: false, next: false, ctl: false, goto: false }), shows: [sh(2)] });
    T("schermo remoto bloccato: con le anteprime spente restano solo le note", hasF("remlock") && !hasF("showcur") && !hasF("shownext"));
    F.send({ control: true, ui: U({ lock: false }), shows: [sh(2)] });
    T("schermo remoto sbloccato dal PC: torna la vista normale (scelte del telefono)", !hasF("remlock") && hasF("nostyle") && !hasF("showcur"));
  }
  // lucchetto
  E.$("slock").onclick();
  T("v3: lucchetto -> schermo bloccato", hasE("locked"));
  E.run("LOCK_HOLD_MS = 60");
  E.$("sunlock").onpointerdown({ preventDefault() {} });
  E.$("sunlock").onpointerup();
  T("v3: rilasciare prima del tempo non sblocca", hasE("locked") && !E.$("sunlock")._cls.has("holding"));
  E.$("sunlock").onpointerdown({ preventDefault() {} });
  T("v3: tenendo premuto parte l'avanzamento", E.$("sunlock")._cls.has("holding") && hasE("locked"));
  setTimeout(() => {
    T("v3: dopo il tempo di pressione lo schermo si sblocca", !hasE("locked") && !E.$("sunlock")._cls.has("holding"));
    E.$("lockb").onclick();
    T("v3: lucchetto anche dalla barra", hasE("locked"));
    E.$("sunlock").onpointerdown({ preventDefault() {} }); E.$("sunlock").onpointerleave();
    T("v3: uscire dal pulsante annulla lo sblocco", hasE("locked"));
    E.run("LOCK_HOLD_MS = 5000"); E.$("sunlock").onpointerdown({ preventDefault() {} }); E.$("sunlock").onpointerup();
    // schermo sempre acceso: parte al primo tocco
    T("v3: video silenzioso anti-standby creato", E.vids().length === 1 && E.vids()[0].played > 0);
    finish3();
  }, 150);
  function finish3() {
    D.send({ control: true, shows: [sh(2)] });   // vecchio formato senza 'ui': tutto disponibile
    T("v3: stato senza 'ui' = tutto disponibile", !hasD("nostyle") && !hasD("noctl") && D.$("styleb").style.display === "");
    console.log(fails ? "CI SONO " + fails + " ERRORI" : "TUTTO OK"); process.exit(fails ? 1 : 0);
  }
  return;
}
}
