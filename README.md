# MEFISTOFELE DeckWhisperer

**Remote notes & control for PowerPoint**

🇬🇧 Reads the **speaker notes** from PowerPoint while the slide show is running and shows them in real time on a web page
you open on your phone (QR code), with text-style controls, next/previous slide buttons and current/next slide previews.
→ **[Manual in English](docs/MANUAL-en.md)**

🇮🇹 Legge le **note del relatore** da PowerPoint mentre la presentazione è in corso e le mostra in tempo reale su una pagina
web da aprire sul telefono (QR code), con comandi per lo stile del testo, avanti/indietro e anteprime della slide in onda e successiva.
→ **[Manuale in italiano](docs/MANUAL-it.md)**

Manuals / Manuali: [English](docs/MANUAL-en.md) · [Italiano](docs/MANUAL-it.md) · [Français](docs/MANUAL-fr.md) · [Español](docs/MANUAL-es.md) · [العربية](docs/MANUAL-ar.md) · [中文](docs/MANUAL-zh.md)

Ideato e sviluppato da [www.andreabattaglioli.it](https://www.andreabattaglioli.it) ·
Powered by [MyWebTvLive.com](https://www.mywebtvlive.com) · Made in Bolognina

## Download / Scarica
**Releases** → `mefistofele.exe` (Windows 10/11 + PowerPoint) e `mefistofele-mac.zip` (macOS 12+ + PowerPoint per Mac). Ogni release contiene anche `SHA256SUMS.txt` e i sei manuali `MANUAL-xx.md`.
Ultima versione in sviluppo: **Actions → Compila**, l'esecuzione più recente, in fondo **mefistofele-windows** e **mefistofele-mac**.

## Novità della 4.0 / What's new in 4.0
- **Accesso remoto / Remote access**: per i telefoni fuori dalla rete (rete dati, altro Wi-Fi). Spento di default, si attiva dalla pagina di gestione; link privato e QR che cambiano a ogni attivazione, PIN sempre richiesto per i comandi, nessuna porta da aprire sul router. Passa dal relay `remote.mefistofele.com` (`relay/`, vedi `docs/RELAY.md`) / For phones outside the network. Off by default, switched on from the management page; private link and QR that change at every activation, PIN always required for controls, nothing to open on the router. Goes through the relay `remote.mefistofele.com`.
- Senza PIN i telefoni remoti possono solo leggere le note; dopo 5 PIN errati blocco di 5 minuti / Without a PIN remote phones can only read the notes; 5 wrong PINs lock access for 5 minutes.
- Note e anteprime passano dal relay cifrate (HTTPS) e non vengono salvate; il relay inoltra solo i percorsi consentiti / Notes and previews cross the relay encrypted (HTTPS) and are not stored; the relay only forwards allowed paths.
- Opzioni `--remote` e `--relay` / `--remote` and `--relay` options.

### Differenze tra la 3.2 e la 4.0 / 3.2 vs 4.0
| | 3.2 | 4.0 |
|---|---|---|
| Telefoni nella stessa rete del computer / Phones on the computer's network | Sì / Yes | Sì / Yes |
| Telefoni fuori dalla rete / Phones outside the network | No | Sì, accesso remoto facoltativo / Yes, optional remote access |
| Cosa esce dalla tua rete / What leaves your network | Niente / Nothing | Niente, finché l'accesso remoto è spento; se attivo, note e anteprime passano dal relay cifrate / Nothing while remote access is off; when on, notes and previews cross the relay, encrypted |
| Impostazioni su router/firewall / Router or firewall settings | Nessuna / None | Nessuna (il computer si collega in uscita) / None (the computer connects outward) |
| PIN | Facoltativo / Optional | Facoltativo in rete locale; obbligatorio per i comandi da fuori, senza PIN solo lettura / Optional on the local network; required for controls from outside, read-only without it |
| Server di terze parti / Third-party server | No | Solo se attivi l'accesso remoto: relay MEFISTOFELE / Only if you switch remote access on: the MEFISTOFELE relay |
| Aggiornamento / Update | – | Si scarica di nuovo il file, le scelte restano / Download the file again, your choices are kept |

La 3.2 resta scaricabile dalla pagina [Releases](https://github.com/battacharlie/PPT-Notes/releases/tag/v3.2) per chi non vuole alcun collegamento verso l'esterno / 3.2 remains available on the Releases page for anyone who wants no outbound connection at all.

## Novità della 3.2 / What's new in 3.2
- Pagina remota: con la barra aperta è sempre visibile per intero e il resto della pagina scende / Remote page: the open toolbar is fully visible and the rest of the page moves down.

## Novità della 3.1 / What's new in 3.1
- **Lingue / Languages**: English (predefinita / default), Italiano, Français, Español, العربية, 中文. Si sceglie nella pagina di gestione e vale per tutti i telefoni e per la guida / chosen on the management page, applies to all phones and to the guide.
- **Manuale in sei lingue / Manual in six languages** (`docs/MANUAL-en|it|fr|es|ar|zh.md`), anche dentro il programma (pulsante Guida) / also inside the program (Guide button).
- **Versione Mac / Mac version**: PowerPoint per Mac tramite AppleScript (`mefistofele-mac.zip`) / PowerPoint for Mac via AppleScript.

## Novità della 3.0 / What's new in 3.0
- Nome **MEFISTOFELE DeckWhisperer** · sottotitolo *Remote notes & control for PowerPoint*.
- Flag rinominati (**Barra degli strumenti**, **Stile Testo**, **Slide in Onda**, **Slide Successiva**, **Vai alla Slide**, **Navigazione**), nuovo flag **Blocca schermo Remoto**, file **mefistofele.exe**, logo MEFISTOFELE / Renamed flags, new **Lock remote screen** flag, **mefistofele.exe** file, MEFISTOFELE logo.
- Pagina di gestione: scelta della **scheda di rete** del QR, flag riordinati, nuovo flag **Vai direttamente alla slide**, pulsante **Guida** / Start panel: **network adapter** for the QR, reordered flags, new **Go to slide** flag, **Guide** button.
- Telefono: **niente standby**, comandi stile **+50%**, dimensione con campo numerico e **a− / A+**, **striscia in alto** con Barra / Schermo / Navigazione / **lucchetto** (sblocco tenendo premuto 5 secondi) / Phone: **no screen sleep**, **+50%** style controls, size box with **a− / A+**, **top strip** with a **screen lock** (hold 5 seconds to unlock).

## Sviluppo / Development
- Sorgente: `windows/deckwhisperer.py` (Python + pywin32 su Windows / AppleScript su Mac + qrcode). Traduzioni in `windows/assets/i18n.json`. Icone in `windows/assets/` (generate da `tools/make_icons.py` a partire da `deckwhisperer-logo.png`).
- Build locale: `windows/crea_deckwhisperer_exe.bat`.
- Test (senza PowerPoint): `python tools/test_deckwhisperer.py` (richiede Node.js) e `python tools/test_remote.py` (accesso remoto: relay + programma).
- Build automatica: GitHub Actions (`.github/workflows/build-windows.yml`, "Compila") a ogni push, con collaudo dell'exe in modalità demo.

**Pubblicare una release**: scrivi il nuovo numero nel file `VERSION` (es. `2.2`) e porta le modifiche su `main`
con una Pull Request: la compilazione automatica crea da sola la release `v2.2`.

## Da fare / To do
- App Android (non prevista per ora / not planned for now).
- Versione Mac: provata solo con AppleScript simulato nei test; da verificare su un Mac reale / Mac version: tested only with simulated AppleScript, to be verified on a real Mac.
