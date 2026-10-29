# DeckWhisperer – Istruzioni (italiano)

**Remote presenter notes & presentation control**
English version: [INSTRUCTIONS.md](INSTRUCTIONS.md)

DeckWhisperer legge le **note del relatore** da PowerPoint mentre la presentazione è in corso e le mostra, in tempo reale,
su una pagina web raggiungibile dal telefono (o da qualsiasi dispositivo della stessa rete). Dal telefono si possono anche
far avanzare le slide e vedere la slide in onda e quella successiva.

## Cosa serve
- Un PC **Windows 10/11** con **PowerPoint** installato (versione desktop).
- Il PC e il telefono collegati alla **stessa rete** (Wi-Fi o cavo).

## Avvio in 4 passi
1. Scarica `DeckWhisperer.exe` dalla pagina [Releases](https://github.com/battacharlie/PPT-Notes/releases) (ultima versione).
2. Fai doppio clic. Se Windows mostra "Windows ha protetto il PC": **Ulteriori informazioni → Esegui comunque**.
   Al primo avvio consenti l'accesso alle **Reti private**.
3. Si apre nel browser del PC la **pagina di avvio** con il QR code. Inquadralo con il telefono.
4. In PowerPoint avvia la presentazione (**Presentazione → Dall'inizio**): le note compaiono sul telefono e seguono le slide.

Il programma resta in una finestra nera: **chiudila per fermarlo**.

## La pagina di avvio (sul PC)
Indirizzo: `http://localhost:8080/start` (si apre da sola). Qui scegli **cosa vedranno i telefoni**:

| Casella | Effetto sul telefono |
|---|---|
| Stile del testo | Font, dimensione, colori, interlinea, spaziatura |
| Comandi slide | Pulsanti Avanti / Indietro / Vai a slide |
| Slide in onda | Immagine della slide che il pubblico sta vedendo |
| Slide successiva | Immagine della slide che arriva dopo |
| Barra dei comandi in alto | Mostra o nasconde tutta la barra dei pulsanti (restano note e scelta presentazione) |
| Modalità tutto schermo | Compare il pulsante *Schermo intero* e al primo tocco la pagina va a tutto schermo (non disponibile su iPhone) |

Puoi anche impostare un **PIN** per i comandi: il telefono lo chiede alla prima pressione di Avanti/Indietro.
Le scelte si **salvano** (file `DeckWhisperer.json` accanto all'exe) e arrivano **subito** ai telefoni già collegati.
La pagina di avvio si apre **solo dal PC** dove gira il programma.

## La pagina sul telefono
Ogni funzione ha il suo colore. Il pulsante **pieno** = acceso, **con solo il bordo** = spento.

| Colore | Pulsante | Cosa fa |
|---|---|---|
| Viola | Stile testo | Mostra/nasconde i comandi del testo |
| Verde | Comandi slide | Mostra/nasconde Avanti / Indietro |
| Arancione | Slide in onda | Mostra/nasconde l'anteprima della slide in onda |
| Blu | Slide successiva | Mostra/nasconde l'anteprima della slide successiva |
| Turchese | Schermo intero | Entra/esce dal tutto schermo |
| Grigio | QR / Info | QR per collegare un altro telefono / versione e crediti |
| Giallo | Nascondi barra | Nasconde la barra (si riapre col pulsante *Barra*) |

In basso: **Indietro** (arancione) e **Avanti** (verde) grandi, più il campo **Vai** per saltare a un numero di slide.
Se ci sono più presentazioni aperte, scegli quale seguire dal menu **Presentazione**.
Ogni telefono ricorda le sue scelte (stile del testo, cosa mostrare).

## Opzioni da riga di comando
`DeckWhisperer.exe --port 8080 --pin 1234 --no-control --no-browser --demo`

- `--port` porta del server (predefinita 8080)
- `--pin` PIN per i comandi
- `--no-control` disattiva del tutto i comandi dal telefono
- `--no-browser` non apre la pagina di avvio sul PC
- `--demo` prova senza PowerPoint

## Se qualcosa non funziona
- **Il telefono non si collega**: stessa rete? Firewall di Windows: consenti le *Reti private* (e pubbliche se il Wi-Fi è visto come pubblico).
- **Le note non compaiono**: la presentazione deve essere **in esecuzione** (a schermo intero), non solo aperta.
- **Le anteprime non compaiono**: apri `http://localhost:8080/diagnostica` sul PC, premi *Prova le anteprime ora* e *Copia tutto*.
- Gli errori sono scritti anche in `DeckWhisperer.log`, accanto all'exe.

---
Ideato e sviluppato da [www.andreabattaglioli.it](https://www.andreabattaglioli.it) ·
Powered by [MyWebTvLive.com](https://www.mywebtvlive.com) · Made in Bolognina
