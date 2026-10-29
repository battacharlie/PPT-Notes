# DeckWhisperer

**Remote presenter notes & presentation control**

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

## Novità della 2.1 / What's new in 2.1
- Nuovo nome e nuovo logo ovunque (icona del file `.exe`, finestra del programma, pagina web) / New name and logo everywhere (exe icon, program window, web page).
- Pannello di avvio con caselle per la **barra dei comandi** e la **modalità tutto schermo** sui telefoni / Start panel with options for the **control bar** and **full-screen mode** on phones.
- Pagina del telefono con pulsanti **colorati**, più grandi e facili da usare / Phone page with **colour-coded**, bigger buttons.
- Istruzioni in italiano e inglese / Instructions in Italian and English.

## Sviluppo / Development
- Sorgente: `windows/deckwhisperer.py` (Python + pywin32 + qrcode). Icone in `windows/assets/` (generate da `tools/make_icons.py` a partire da `deckwhisperer-logo.png`).
- Build locale: `windows/crea_deckwhisperer_exe.bat`.
- Test (senza PowerPoint): `python tools/test_deckwhisperer.py` (richiede Node.js).
- Build automatica: GitHub Actions (`.github/workflows/build-windows.yml`, "Compila") a ogni push, con collaudo dell'exe in modalità demo.

**Pubblicare una release**: scrivi il nuovo numero nel file `VERSION` (es. `2.2`) e porta le modifiche su `main`
con una Pull Request: la compilazione automatica crea da sola la release `v2.2`.

## Da fare / To do
- Mac (lettura via AppleScript) e app Android (lettore QR, scelta presentazione, schermo sempre acceso, comandi).
