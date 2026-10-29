# DeckWhisperer – Instructions (English)

**Remote presenter notes & presentation control**
Versione italiana: [ISTRUZIONI.md](ISTRUZIONI.md)

DeckWhisperer reads the **speaker notes** from PowerPoint while your presentation is running and shows them in real time
on a web page you open on your phone (or any device on the same network). From the phone you can also move the slides
forward and back, and see the current and the next slide.

## Requirements
- A **Windows 10/11** PC with desktop **PowerPoint** installed.
- The PC and the phone on the **same network** (Wi-Fi or cable).

## Quick start
1. Download `DeckWhisperer.exe` from the [Releases](https://github.com/battacharlie/PPT-Notes/releases) page (latest version).
2. Double-click it. If Windows shows "Windows protected your PC": **More info → Run anyway**.
   On first start, allow access on **Private networks**.
3. The **start page** opens in the PC browser with a QR code. Scan it with your phone.
4. In PowerPoint start the slide show (**Slide Show → From Beginning**): the notes appear on the phone and follow the slides.

The program runs in a black window: **close it to stop the program**.

## The start page (on the PC)
Address: `http://localhost:8080/start` (opens by itself). Here you choose **what the phones will show**:

| Checkbox | Effect on the phone |
|---|---|
| Text style | Font, size, colours, line spacing, letter spacing |
| Slide controls | Next / Previous / Go-to-slide buttons |
| Current slide | Picture of the slide the audience is seeing |
| Next slide | Picture of the upcoming slide |
| Top control bar | Shows or hides the whole button bar (notes and presentation picker stay) |
| Full-screen mode | The *Full screen* button appears and the page goes full screen on the first tap (not available on iPhone) |

You can also set a **PIN** for the controls: the phone asks for it the first time you press Next/Previous.
Your choices are **saved** (`DeckWhisperer.json` next to the exe) and reach connected phones **instantly**.
The start page opens **only from the PC** running the program.

## The page on the phone
Every function has its own colour. A **filled** button = on, an **outline-only** button = off.

| Colour | Button | What it does |
|---|---|---|
| Purple | Stile testo (Text style) | Show/hide the text controls |
| Green | Comandi slide (Slide controls) | Show/hide Next / Previous |
| Orange | Slide in onda (Current slide) | Show/hide the current-slide preview |
| Blue | Slide successiva (Next slide) | Show/hide the next-slide preview |
| Teal | Schermo intero (Full screen) | Enter/leave full screen |
| Grey | QR / Info | QR to connect another phone / version and credits |
| Yellow | Nascondi barra (Hide bar) | Hides the bar (bring it back with the *Barra* button) |

At the bottom: big **Indietro** (Previous, orange) and **Avanti** (Next, green) buttons, plus the **Vai** (Go) box to jump to a slide number.
If several presentations are open, pick the one to follow from the **Presentazione** menu.
Each phone remembers its own choices (text style, what to show).

## Command-line options
`DeckWhisperer.exe --port 8080 --pin 1234 --no-control --no-browser --demo`

- `--port` server port (default 8080)
- `--pin` PIN for the controls
- `--no-control` turns the phone controls off completely
- `--no-browser` does not open the start page on the PC
- `--demo` try it without PowerPoint

## Troubleshooting
- **The phone can't connect**: same network? Windows Firewall: allow *Private networks* (and public ones if your Wi-Fi is seen as public).
- **No notes shown**: the slide show must be **running** (full screen), not just open.
- **Previews don't show**: open `http://localhost:8080/diagnostica` on the PC, press *Prova le anteprime ora* and *Copia tutto*.
- Errors are also written to `DeckWhisperer.log` next to the exe.

---
Created and developed by [www.andreabattaglioli.it](https://www.andreabattaglioli.it) ·
Powered by [MyWebTvLive.com](https://www.mywebtvlive.com) · Made in Bolognina
