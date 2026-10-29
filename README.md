# reMNote

Legge le **note del relatore** da PowerPoint mentre la presentazione è in corso e le mostra, in tempo reale,
su una pagina web raggiungibile da qualsiasi dispositivo della rete locale (smartphone, tablet, altri PC).

Ideato e sviluppato da [www.andreabattaglioli.it](https://www.andreabattaglioli.it) ·
Powered by [MyWebTvLive.com](https://www.mywebtvlive.com) · Made in Bolognina

## Funzioni (Windows)
- Note sincronizzate con la slide corrente (aggiornamento ~10 volte al secondo).
- QR code in console e su `http://localhost:8080/qr` per aprire le note dallo smartphone.
- Più presentazioni aperte insieme: ogni dispositivo sceglie quale seguire.
- Barra stile: font, dimensione, colore testo, colore sfondo, interlinea, spaziatura caratteri (salvata per dispositivo).
- Comandi avanti / indietro / vai a slide dal telefono, con PIN opzionale (`--pin 1234`) o disattivabili (`--no-control`).

## Uso
**Versioni ufficiali**: nella pagina del progetto su GitHub, sezione **Releases** (a destra), scarica `reMNote.exe`.
Ogni release contiene anche `SHA256SUMS.txt` per verificare il file.

**Ultima versione in sviluppo**: **Actions → Compila**, l'esecuzione più recente, in fondo **reMNote-windows**.

Avvia `reMNote.exe` sul PC dove gira PowerPoint e inquadra il QR con lo smartphone (stessa rete Wi-Fi).
Se compare "Windows ha protetto il PC": *Ulteriori informazioni → Esegui comunque*.
Alla prima esecuzione Windows chiede l'accesso alla rete: consenti le *Reti private*.

Opzioni: `--port 8080`, `--pin 1234`, `--no-control`, `--no-browser`, `--demo` (prova senza PowerPoint).

## Sviluppo
- Sorgente: `windows/remnote.py` (Python + pywin32 + qrcode). Icone in `windows/assets/` (generate da `tools/make_icons.py` a partire da `remnote-logo.png`).
- Build locale: `windows/crea_remnote_exe.bat`.
- Build automatica: GitHub Actions (`.github/workflows/build-windows.yml`, "Compila") a ogni push, con collaudo dell'exe in modalità demo.

**Pubblicare una release**: scrivi il nuovo numero nel file `VERSION` (es. `1.1`) e porta le modifiche su `main`
(con una Pull Request): la compilazione automatica crea da sola la release `v1.1` con `reMNote.exe` e `SHA256SUMS.txt`.

## Da fare
- Mac (richiede lettura via AppleScript) e app Android (lettore QR, scelta presentazione, schermo sempre acceso, comandi).
