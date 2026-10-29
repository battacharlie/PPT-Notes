# MEFISTOFELE DeckWhisperer

**Remote notes & control for PowerPoint**

🇬🇧 Reads the **speaker notes** from PowerPoint while the slide show is running and shows them in real time on a web page
you open on your phone (QR code), with text-style controls, next/previous slide buttons and current/next slide previews.
→ **[Instructions in English](docs/INSTRUCTIONS.md)**

🇮🇹 Legge le **note del relatore** da PowerPoint mentre la presentazione è in corso e le mostra in tempo reale su una pagina
web da aprire sul telefono (QR code), con comandi per lo stile del testo, avanti/indietro e anteprime della slide in onda e successiva.
→ **[Istruzioni in italiano](docs/ISTRUZIONI.md)**

Ideato e sviluppato da [www.andreabattaglioli.it](https://www.andreabattaglioli.it) ·
Powered by [MyWebTvLive.com](https://www.mywebtvlive.com) · Made in Bolognina

## Download / Scarica
**Releases** → `DeckWhisperer.exe` (Windows 10/11 + PowerPoint). Ogni release contiene anche `SHA256SUMS.txt`, `INSTRUCTIONS.md` e `ISTRUZIONI.md`.
Ultima versione in sviluppo: **Actions → Compila**, l'esecuzione più recente, in fondo **DeckWhisperer-windows**.

## Novità della 3.0 / What's new in 3.0
- Nome **MEFISTOFELE DeckWhisperer** · sottotitolo *Remote notes & control for PowerPoint*.
- Pannello di avvio: scelta della **scheda di rete** del QR, flag riordinati, nuovo flag **Vai direttamente alla slide**, pulsante **Guida** / Start panel: **network adapter** for the QR, reordered flags, new **Go to slide** flag, **Guide** button.
- Telefono: **niente standby**, comandi stile **+50%**, dimensione con campo numerico e **a− / A+**, **striscia in alto** con Barra / Schermo / Comandi / **lucchetto** (sblocco tenendo premuto 5 secondi) / Phone: **no screen sleep**, **+50%** style controls, size box with **a− / A+**, **top strip** with a **screen lock** (hold 5 seconds to unlock).

## Sviluppo / Development
- Sorgente: `windows/deckwhisperer.py` (Python + pywin32 + qrcode). Icone in `windows/assets/` (generate da `tools/make_icons.py` a partire da `deckwhisperer-logo.png`).
- Build locale: `windows/crea_deckwhisperer_exe.bat`.
- Test (senza PowerPoint): `python tools/test_deckwhisperer.py` (richiede Node.js).
- Build automatica: GitHub Actions (`.github/workflows/build-windows.yml`, "Compila") a ogni push, con collaudo dell'exe in modalità demo.

**Pubblicare una release**: scrivi il nuovo numero nel file `VERSION` (es. `2.2`) e porta le modifiche su `main`
con una Pull Request: la compilazione automatica crea da sola la release `v2.2`.

## Da fare / To do
- Mac (lettura via AppleScript) e app Android (lettore QR, scelta presentazione, schermo sempre acceso, comandi).
