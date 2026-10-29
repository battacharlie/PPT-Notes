# MEFISTOFELE DeckWhisperer – Instructions (English)

**Remote notes & control for PowerPoint**
Versione italiana: [ISTRUZIONI.md](ISTRUZIONI.md)

In these instructions the **management page** is the page that opens on the PC running PowerPoint; the **remote page** is the one you see on the phone.

MEFISTOFELE DeckWhisperer reads the **speaker notes** from PowerPoint while your presentation is running and shows them in real time
on a web page you open on your phone (or any device on the same network). From the phone you can also move the slides
forward and back, jump to a slide, and see the current and the next slide.

## Requirements
- A **Windows 10/11** PC with desktop **PowerPoint** installed.
- The PC and the phone on the **same network** (Wi-Fi or cable).

## Quick start
1. Download `mefistofele.exe` from the [Releases](https://github.com/battacharlie/PPT-Notes/releases) page (latest version).
2. Double-click it. If Windows shows "Windows protected your PC": **More info → Run anyway**.
   On first start, allow access on **Private networks**.
3. The **management page** opens in the PC browser with a QR code. Scan it with your phone.
4. In PowerPoint start the slide show (**Slide Show → From Beginning**): the notes appear on the phone and follow the slides.

The program runs in a black window: **close it to stop the program**.

## The management page (on the PC)
Address: `http://localhost:8080/start` (opens by itself). The **Guida** (Guide) button at the top right opens these instructions.

**Network adapter of the QR.** If the PC has several network adapters (for example Wi-Fi and a production network), pick from the
**Scheda di rete del QR** menu the one your phones are connected to: the QR and the address change immediately.
*Automatica* (automatic) lets the program choose.

Then choose **what the phones will show**. The list items, in order:

| Checkbox | Effect on the phone |
|---|---|
| Barra degli strumenti (Toolbar) | The phone shows the bar of colour-coded buttons, with the strip and the padlock. **If it is off**, the phone shows only the items chosen below and **nothing can be changed** |
| Stile Testo (Text style) | Font, size, colours, line spacing, letter spacing |
| Slide in Onda (Current slide) | Picture of the slide the audience is seeing |
| Slide Successiva (Next slide) | Picture of the upcoming slide |
| Vai alla Slide (Go to slide) | Number box and **Vai** button to jump to a slide |
| Navigazione (Navigation) | Big **Avanti** (Next) and **Indietro** (Previous) buttons |
| Blocca schermo Remoto (Lock remote screen) | Removes **all controls** from the phone: only the notes stay, which can just be **scrolled** if longer than the screen, plus the chosen previews. Commands are no longer accepted |

You can also set a **PIN** for the controls: the phone asks for it the first time you press Next/Previous or Go.
Your choices are **saved** (`mefistofele.json` next to the exe) and reach connected phones **instantly**.
The management page opens **only from the PC** running the program.

## The remote page (on the phone)
Every function has its own colour. A **filled** button = on, an **outline-only** button = off.

| Colour | Button | What it does |
|---|---|---|
| Purple | Stile testo (Text style) | Show/hide the text controls |
| Green | Navigazione (Navigation) | Show/hide Next / Previous and Go to slide |
| Orange | Slide in onda (Current slide) | Show/hide the current-slide preview |
| Blue | Slide successiva (Next slide) | Show/hide the next-slide preview |
| Teal | Schermo intero (Full screen) | Enter/leave full screen (not available on iPhone) |
| Red | Blocca schermo (Lock screen) | Locks the phone against accidental touches |
| Grey | QR / Info | QR to connect another phone / version and credits |
| Yellow | Nascondi barra (Hide bar) | Closes the bar |

**Text size**: type the number in the box, or use **a−** (smaller) and **A+** (bigger).

**Top strip.** When you close the bar a strip stays with four buttons: **Barra** (reopens the bar),
**Schermo** (full screen on/off), **Navigazione** (show/hide Next, Previous and Go) and **Blocca** (lock).

**Screen lock.** With the padlock (available when the Toolbar is on) the screen ignores every touch (the notes keep updating).
To **unlock, press and hold the red button at the top for 5 seconds**: a bar fills up and then the screen is released.
If you let go earlier, it stays locked.

**The phone does not go to sleep** while the page is open. The first tap on the page turns the protection on;
if the phone still turns off, check that *battery saver* is not active.

At the bottom: big **Indietro** (Previous, orange) and **Avanti** (Next, green) buttons, plus the **Vai** (Go) box to jump to a slide number.
If several presentations are open, pick the one to follow from the **Presentazione** menu (not available with *Lock remote screen*).
Each phone remembers its own choices (text style, what to show).

## Command-line options
`mefistofele.exe --port 8080 --pin 1234 --no-control --no-browser --demo`

- `--port` server port (default 8080)
- `--pin` PIN for the controls
- `--no-control` turns the phone controls off completely
- `--no-browser` does not open the management page on the PC
- `--demo` try it without PowerPoint

## Troubleshooting
- **The phone can't connect**: same network? If the PC has several adapters, pick the right one on the management page. Windows Firewall: allow *Private networks* (and public ones if your Wi-Fi is seen as public).
- **No notes shown**: the slide show must be **running** (full screen), not just open.
- **Previews don't show**: open `http://localhost:8080/diagnostica` on the PC, press *Prova le anteprime ora* and *Copia tutto*.
- Errors are also written to `mefistofele.log` next to the exe.

---
Created and developed by [www.andreabattaglioli.it](https://www.andreabattaglioli.it) ·
Powered by [MyWebTvLive.com](https://www.mywebtvlive.com) · Made in Bolognina
