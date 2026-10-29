# MEFISTOFELE DeckWhisperer – Manuale (Italiano)

**Remote notes & control for PowerPoint**
[English](MANUAL-en.md) · [Italiano](MANUAL-it.md) · [Français](MANUAL-fr.md) · [Español](MANUAL-es.md) · [العربية](MANUAL-ar.md) · [中文](MANUAL-zh.md)

In questo manuale la **pagina di gestione** è la pagina che si apre sul computer con PowerPoint; la **pagina remota** è quella che vedi sul telefono.

MEFISTOFELE DeckWhisperer legge le **note del relatore** di PowerPoint mentre la presentazione è in corso e le mostra in tempo reale
su una pagina web che apri dal telefono (o da qualsiasi dispositivo nella stessa rete). Dal telefono puoi anche far avanzare e tornare indietro le slide,
saltare a una slide e vedere la slide in onda e la successiva.

## Requisiti
- **Windows 10/11** con **PowerPoint** desktop, oppure un **Mac** (macOS 12 o successivo) con **PowerPoint per Mac**.
- Computer e telefono nella **stessa rete** (Wi-Fi o cavo).

## Avvio in 4 passi
**Windows**
1. Scarica `mefistofele.exe` dalla pagina [Releases](https://github.com/battacharlie/PPT-Notes/releases) (ultima versione).
2. Doppio clic. Se Windows mostra "Windows ha protetto il PC": **Ulteriori informazioni → Esegui comunque**. Al primo avvio consenti l'accesso alle **Reti private**.

**Mac**
1. Scarica `mefistofele-mac.zip` dalla pagina [Releases](https://github.com/battacharlie/PPT-Notes/releases) e decomprimilo.
2. Apri **Terminale**, scrivi `chmod +x ` (con lo spazio), trascina il file `mefistofele` nella finestra e premi Invio. Poi ripeti senza `chmod +x ` per avviarlo.
   Se macOS lo blocca ("non può essere aperto perché lo sviluppatore non è verificato"): esegui `xattr -d com.apple.quarantine ` seguito dal file, oppure **Impostazioni di Sistema → Privacy e sicurezza → Apri comunque**.
3. La prima volta macOS chiede se **il Terminale può controllare Microsoft PowerPoint**: premi **OK** (si cambia poi in **Impostazioni di Sistema → Privacy e sicurezza → Automazione**).

**Poi (su entrambi)**
- Nel browser si apre la **pagina di gestione** con un QR code. Inquadralo con il telefono.
- In PowerPoint avvia la presentazione (**Presentazione → Dall'inizio**): le note compaiono sul telefono e seguono le slide.

Il programma resta in una finestra nera (il Terminale su Mac): **chiudila per fermare il programma**.

## La pagina di gestione (sul computer)
Indirizzo: `http://localhost:8080/start` (si apre da sola). Il pulsante **Guida** in alto a destra apre questo manuale nella lingua scelta.

**Lingua.** Scegli la lingua dal menu **Lingua**: English (predefinita), Italiano, Français, Español, العربية o 中文.
Vale per la pagina di gestione, per la pagina remota su **tutti i telefoni** (seguono la pagina di gestione) e per la guida. La scelta viene salvata.

**Scheda di rete del QR.** Se il computer ha più schede di rete (per esempio Wi-Fi e rete di produzione), scegli dal menu
**Scheda di rete del QR** quella a cui sono collegati i telefoni: QR e indirizzo cambiano subito. *Automatica* lascia scegliere al programma.

Poi scegli **cosa mostrano i telefoni**. Le voci, in ordine:

| Casella | Effetto sul telefono |
|---|---|
| Barra degli strumenti | Il telefono mostra la barra di pulsanti colorati, con la striscia e il lucchetto. **Se è spenta**, il telefono mostra solo le voci scelte qui sotto e **non si può modificare nulla** |
| Stile Testo | Font, dimensione, colori, interlinea, spaziatura |
| Slide in Onda | Immagine della slide che vede il pubblico |
| Slide Successiva | Immagine della slide che segue |
| Vai alla Slide | Casella numerica e pulsante **Vai** per saltare a una slide |
| Navigazione | Grandi pulsanti **Avanti** e **Indietro** |
| Blocca schermo Remoto | Toglie **tutti i comandi** dal telefono: restano solo le note, che si possono solo **scorrere** se superano lo schermo, più le anteprime scelte. I comandi non vengono più accettati |

Puoi anche impostare un **PIN** per i comandi: il telefono lo chiede la prima volta che premi Avanti/Indietro o Vai.
Le scelte vengono **salvate** (`mefistofele.json` accanto al programma) e arrivano ai telefoni collegati **subito**.
La pagina di gestione si apre **solo dal computer** su cui gira il programma.

## La pagina remota (sul telefono)
Ogni funzione ha un suo colore. Pulsante **pieno** = acceso, pulsante **solo contorno** = spento.

| Colore | Pulsante | Cosa fa |
|---|---|---|
| Viola | Stile Testo | Mostra/nasconde i comandi del testo |
| Verde | Navigazione | Mostra/nasconde Avanti / Indietro e Vai alla slide |
| Arancione | Slide in Onda | Mostra/nasconde l'anteprima della slide in onda |
| Blu | Slide Successiva | Mostra/nasconde l'anteprima della successiva |
| Turchese | Schermo intero | Entra/esce dallo schermo intero (non disponibile su iPhone) |
| Rosso | Blocca schermo | Blocca il telefono contro i tocchi accidentali |
| Grigio | QR / Info | QR per collegare un altro telefono / versione, guida e crediti |
| Giallo | Nascondi barra | Chiude la barra |

**Dimensione del testo**: scrivi il numero nella casella, oppure usa **a−** (più piccolo) e **A+** (più grande).

**Striscia in alto.** Quando chiudi la barra resta una striscia con quattro pulsanti: **Barra** (riapre la barra),
**Schermo** (schermo intero), **Navigazione** (mostra/nasconde Avanti, Indietro e Vai) e **Blocca**.

**Blocco schermo.** Con il lucchetto lo schermo ignora ogni tocco (le note continuano ad aggiornarsi).
Per **sbloccare tieni premuto 5 secondi il pulsante rosso in alto**: una barra si riempie e poi lo schermo si libera. Se lasci prima, resta bloccato.

**Il telefono non va in standby** finché la pagina è aperta. Il primo tocco sulla pagina attiva la protezione;
se il telefono si spegne lo stesso, controlla che il *risparmio energetico* non sia attivo.

In basso: grandi pulsanti **Indietro** e **Avanti**, più la casella **Vai** per saltare a un numero di slide.
Se sono aperte più presentazioni, scegli quella da seguire dal menu **Presentazione** (non disponibile con *Blocca schermo Remoto*).
Ogni telefono ricorda le proprie scelte (stile del testo, cosa mostrare). La lingua è quella scelta nella pagina di gestione.

## Opzioni da riga di comando
Windows: `mefistofele.exe --port 8080 --pin 1234 --no-control --no-browser --demo`
Mac: `./mefistofele --port 8080 --pin 1234 --no-control --no-browser --demo`

- `--port` porta del server (predefinita 8080)
- `--pin` PIN per i comandi
- `--no-control` spegne del tutto i comandi dal telefono
- `--no-browser` non apre la pagina di gestione sul computer
- `--demo` prova senza PowerPoint

## Se qualcosa non va
- **Il telefono non si collega**: stessa rete? Se il computer ha più schede, scegli quella giusta nella pagina di gestione. Firewall di Windows: consenti le *Reti private* (e quelle pubbliche se il Wi-Fi è visto come pubblico). Su Mac consenti le connessioni in entrata se richiesto.
- **Non compaiono le note**: la presentazione deve essere **in corso** (a schermo intero), non solo aperta.
- **Mac: non succede nulla**: controlla **Impostazioni di Sistema → Privacy e sicurezza → Automazione** e verifica che il Terminale possa controllare Microsoft PowerPoint.
- **Anteprime delle slide**: su Mac dipendono dal fatto che PowerPoint consenta l'esportazione; se falliscono, il motivo compare al posto dell'immagine e note e comandi continuano a funzionare.
- **Anteprime assenti (Windows)**: apri `http://localhost:8080/diagnostica`, premi il pulsante di prova e copia tutto.
- Gli errori vengono scritti anche in `mefistofele.log` accanto al programma (o nella cartella temporanea).

---
Created and developed by [www.andreabattaglioli.it](https://www.andreabattaglioli.it) ·
Powered by [MyWebTvLive.com](https://www.mywebtvlive.com) · Made in Bolognina
