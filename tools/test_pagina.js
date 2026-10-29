const fs = require("fs"), vm = require("vm");
const code = fs.readFileSync(process.argv[2], "utf8");
const els = {}; const store = {}; let es = null;
function El(id) {
  const e = { id, style: {}, attrs: {}, _cls: new Set(), textContent: "", value: "", options: [], className: "", max: 0 };
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
  document: { getElementById: $, body, documentElement: Object.assign(docEl, { requestFullscreen: () => Promise.resolve() }),
    createElement: () => El("opt"), addEventListener() {}, fullscreenElement: null, visibilityState: "visible" },
  navigator: {}, localStorage: { getItem: k => (k in store ? store[k] : null), setItem: (k, v) => { store[k] = String(v); } },
  EventSource: function () { es = this; }, fetch: (u) => Promise.resolve({ ok: true, status: 200, text: () => Promise.resolve(globalThis.__errText || "") }),
  prompt: () => null, setTimeout, clearTimeout, console, encodeURIComponent, parseInt, parseFloat, Number, JSON, Math,
};
vm.createContext(sandbox); vm.runInContext(code, sandbox);
let fails = 0; const T = (n, c) => { if (!c) fails++; console.log((c ? "OK   " : "FAIL ") + n); };
const send = st => es.onmessage({ data: JSON.stringify(st) });
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
console.log(fails ? "CI SONO " + fails + " ERRORI" : "TUTTO OK"); process.exit(fails ? 1 : 0);
}
