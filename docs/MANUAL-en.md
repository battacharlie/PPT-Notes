# MEFISTOFELE DeckWhisperer – Manual (English)

**Remote notes & control for PowerPoint**
[English](MANUAL-en.md) · [Italiano](MANUAL-it.md) · [Français](MANUAL-fr.md) · [Español](MANUAL-es.md) · [العربية](MANUAL-ar.md) · [中文](MANUAL-zh.md)

In this manual the **management page** is the page that opens on the computer running PowerPoint; the **remote page** is the one you see on the phone.

MEFISTOFELE DeckWhisperer reads the **speaker notes** from PowerPoint while your presentation is running and shows them in real time
on a web page you open on your phone (or any device on the same network). From the phone you can also move the slides
forward and back, jump to a slide, and see the current and the next slide.

## Requirements
- **Windows 10/11** with desktop **PowerPoint**, or a **Mac** (macOS 12 or later) with **PowerPoint for Mac**.
- The computer and the phone on the **same network** (Wi-Fi or cable).

## Quick start
**Windows**
1. Download `mefistofele.exe` from the [Releases](https://github.com/battacharlie/PPT-Notes/releases) page (latest version).
2. Double-click it. If Windows shows "Windows protected your PC": **More info → Run anyway**. On first start, allow access on **Private networks**.

**Mac**
1. Download `mefistofele-mac.zip` from the [Releases](https://github.com/battacharlie/PPT-Notes/releases) page and unzip it.
2. Open **Terminal**, type `chmod +x ` (with a space), drag the `mefistofele` file into the window and press Enter. Then do the same without `chmod +x ` to run it.
   If macOS blocks it ("cannot be opened because the developer cannot be verified"): run `xattr -d com.apple.quarantine ` followed by the file, or **System Settings → Privacy & Security → Open Anyway**.
3. The first time, macOS asks whether **Terminal may control Microsoft PowerPoint**: click **OK** (it can be changed later in **System Settings → Privacy & Security → Automation**).

**Then (both systems)**
- The **management page** opens in the browser with a QR code. Scan it with your phone.
- In PowerPoint start the slide show (**Slide Show → From Beginning**): the notes appear on the phone and follow the slides.

The program runs in a black window (Terminal on Mac): **close it to stop the program**.

## The management page (on the computer)
Address: `http://localhost:8080/start` (opens by itself). The **Guide** button at the top right opens this manual in the chosen language.

**Language.** Pick the language from the **Language** menu: English (default), Italiano, Français, Español, العربية or 中文.
It applies to the management page, to the remote page on **all phones** (they follow the management page) and to the guide. The choice is saved.

**Network adapter of the QR.** If the computer has several network adapters (for example Wi-Fi and a production network), pick from the
**Network card of the QR** menu the one your phones are connected to: the QR and the address change immediately. *Automatic* lets the program choose.

Then choose **what the phones will show**. The list items, in order:

| Checkbox | Effect on the phone |
|---|---|
| Toolbar | The phone shows the bar of colour-coded buttons, with the strip and the padlock. **If it is off**, the phone shows only the items chosen below and **nothing can be changed** |
| Text style | Font, size, colours, line spacing, letter spacing |
| Current slide | Picture of the slide the audience is seeing |
| Next slide | Picture of the upcoming slide |
| Go to slide | Number box and **Go** button to jump to a slide |
| Navigation | Big **Next** and **Previous** buttons |
| Lock remote screen | Removes **all controls** from the phone: only the notes stay, which can just be **scrolled** if longer than the screen, plus the chosen previews. Commands are no longer accepted |

You can also set a **PIN** for the controls: the phone asks for it the first time you press Next/Previous or Go.
Your choices are **saved** (`mefistofele.json` next to the program) and reach connected phones **instantly**.
The management page opens **only on the computer** running the program.

## The remote page (on the phone)
Every function has its own colour. A **filled** button = on, an **outline-only** button = off.

| Colour | Button | What it does |
|---|---|---|
| Purple | Text style | Show/hide the text controls |
| Green | Navigation | Show/hide Next / Previous and Go to slide |
| Orange | Current slide | Show/hide the current-slide preview |
| Blue | Next slide | Show/hide the next-slide preview |
| Teal | Full screen | Enter/leave full screen (not available on iPhone) |
| Red | Lock screen | Locks the phone against accidental touches |
| Grey | QR / Info | QR to connect another phone / version, guide and credits |
| Yellow | Hide toolbar | Closes the bar |

**Text size**: type the number in the box, or use **a−** (smaller) and **A+** (bigger).

**Top strip.** When you close the bar a strip stays with four buttons: **Toolbar** (reopens the bar),
**Screen** (full screen on/off), **Navigation** (show/hide Next, Previous and Go) and **Lock**.

**Screen lock.** With the padlock the screen ignores every touch (the notes keep updating).
To **unlock, press and hold the red button at the top for 5 seconds**: a bar fills up and then the screen is released. If you let go earlier, it stays locked.

**The phone does not go to sleep** while the page is open. The first tap on the page turns the protection on;
if the phone still turns off, check that *battery saver* is not active.

At the bottom: big **Previous** and **Next** buttons, plus the **Go** box to jump to a slide number.
If several presentations are open, pick the one to follow from the **Presentation** menu (not available with *Lock remote screen*).
Each phone remembers its own choices (text style, what to show). The language is the one chosen on the management page.

## Remote access (phones outside the network)
For phones that are **not** on the same network as the PC (mobile data, another Wi-Fi). It is **off by default**: switch it on only when you need it.
1. In the management page, in the **Remote access** box, press **Switch on remote access**.
2. A private link (`https://remote.mefistofele.com:8443/s/...`) and a QR code appear. Open the link on the remote phone.
3. To let that phone use **Next/Previous/Go to slide**, set a **PIN** in the management page: without a PIN, remote phones can only read the notes. The PIN is always requested from outside, and after 5 wrong attempts the controls are blocked for 5 minutes.
4. Press **Switch off** (or close the program) to end access immediately. The link changes every time you switch it on.

The PC connects **outward** to the MEFISTOFELE relay server: no router or firewall settings are needed. Notes and slide previews pass through the relay encrypted (HTTPS) and are **not stored**.

## Command-line options
Windows: `mefistofele.exe --port 8080 --pin 1234 --no-control --no-browser --demo`
Mac: `./mefistofele --port 8080 --pin 1234 --no-control --no-browser --demo`

- `--port` server port (default 8080)
- `--pin` PIN for the controls
- `--no-control` turns the phone controls off completely
- `--no-browser` does not open the management page on the computer
- `--demo` try it without PowerPoint
- `--remote` switch on remote access at start
- `--relay` address of the relay server (default `https://remote.mefistofele.com:8443`)

## Troubleshooting
- **The phone can't connect**: same network? If the computer has several adapters, pick the right one on the management page. Windows Firewall: allow *Private networks* (and public ones if your Wi-Fi is seen as public). On Mac allow incoming connections if asked.
- **No notes shown**: the slide show must be **running** (full screen), not just open.
- **Mac: nothing happens**: check **System Settings → Privacy & Security → Automation** and make sure Terminal is allowed to control Microsoft PowerPoint.
- **Slide previews**: on Mac they depend on PowerPoint allowing the export; if they fail, the reason is shown in place of the picture and the notes and controls keep working.
- **Previews don't show (Windows)**: open `http://localhost:8080/diagnostica`, press the test button and copy everything.
- Errors are also written to `mefistofele.log` next to the program (or in the temporary folder).

---
Created and developed by [www.andreabattaglioli.it](https://www.andreabattaglioli.it) ·
Powered by [MyWebTvLive.com](https://www.mywebtvlive.com) · Made in Bolognina
