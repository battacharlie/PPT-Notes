# MEFISTOFELE DeckWhisperer – Istruzioni (italiano)

**Remote notes & control for PowerPoint**
English version: [INSTRUCTIONS.md](INSTRUCTIONS.md)

In queste istruzioni la **pagina di gestione** è la pagina che si apre sul PC dove gira PowerPoint; la **pagina remota** è quella che si vede sul telefono.

MEFISTOFELE DeckWhisperer legge le **note del relatore** da PowerPoint mentre la presentazione è in corso e le mostra, in tempo reale,
su una pagina web raggiungibile dal telefono (o da qualsiasi dispositivo della stessa rete). Dal telefono si possono anche
far avanzare le slide, saltare a una slide e vedere la slide in onda e quella successiva.

## Cosa serve
- Un PC **Windows 10/11** con **PowerPoint** installato (versione desktop).
- Il PC e il telefono collegati alla **stessa rete** (Wi-Fi o cavo).

## Avvio in 4 passi
1. Scarica `mefistofele.exe` dalla pagina [Releases](https://github.com/battacharlie/PPT-Notes/releases) (ultima versione).
2. Fai doppio clic. Se Windows mostra "Windows ha protetto il PC": **Ulteriori informazioni → Esegui comunque**.
   Al primo avvio consenti l'accesso alle **Reti private**.
3. Si apre nel browser del PC la **pagina di gestione** con il QR code. Inquadralo con il telefono.
4. In PowerPoint avvia la presentazione (**Presentazione → Dall'inizio**): le note compaiono sul telefono e seguono le slide.

Il programma resta in una finestra nera: **chiudila per fermarlo**.

## La pagina di gestione (sul PC)
Indirizzo: `http://localhost:8080/start` (si apre da sola). In alto a destra il pulsante **Guida** apre queste istruzioni.

**Scheda di rete del QR.** Se il PC ha più schede di rete (per esempio Wi-Fi e rete di produzione), scegli dal menu
**Scheda di rete del QR** quella a cui sono collegati i telefoni: il QR e l'indirizzo cambiano subito.
*Automatica* lascia scegliere al programma.

Poi scegli **cosa vedranno i telefoni**. Gli elementi della lista, in ordine:

| Casella | Effetto sul telefono |
|---|---|
| Barra degli strumenti | Sul telefono compare la barra dei pulsanti colorati, con la striscia e il lucchetto. **Se è spenta**, sul telefono compaiono solo gli elementi scelti qui sotto e **non si può modificare nulla** |
| Stile Testo | Font, dimensione, colori, interlinea, spaziatura |
| Slide in Onda | Immagine della slide che il pubblico sta vedendo |
| Slide Successiva | Immagine della slide che arriva dopo |
| Vai alla Slide | Campo numero e pulsante **Vai** per saltare a una slide |
| Navigazione | Pulsanti grandi **Avanti** e **Indietro** |
| Blocca schermo Remoto | Toglie **tutti i comandi** dal telefono: restano le note, che si possono solo **scorrere** se più lunghe dello schermo, e le anteprime scelte. I comandi non vengono più accettati |

Puoi anche impostare un **PIN** per i comandi: il telefono lo chiede alla prima pressione di Avanti/Indietro o di Vai.
Le scelte si **salvano** (file `mefistofele.json` accanto all'exe) e arrivano **subito** ai telefoni già collegati.
La pagina di gestione si apre **solo dal PC** dove gira il programma.

## La pagina remota (sul telefono)
Ogni funzione ha il suo colore. Il pulsante **pieno** = acceso, **con solo il bordo** = spento.

| Colore | Pulsante | Cosa fa |
|---|---|---|
| Viola | Stile testo | Mostra/nasconde i comandi del testo |
| Verde | Navigazione | Mostra/nasconde Avanti / Indietro e Vai alla slide |
| Arancione | Slide in onda | Mostra/nasconde l'anteprima della slide in onda |
| Blu | Slide successiva | Mostra/nasconde l'anteprima della slide successiva |
| Turchese | Schermo intero | Entra/esce dal tutto schermo (non disponibile su iPhone) |
| Rosso | Blocca schermo | Blocca il telefono contro i tocchi involontari |
| Grigio | QR / Info | QR per collegare un altro telefono / versione e crediti |
| Giallo | Nascondi barra | Chiude la barra |

**Dimensione del testo**: scrivi il numero nel campo oppure usa **a−** (più piccolo) e **A+** (più grande).

**Striscia in alto.** Quando chiudi la barra resta una striscia con quattro pulsanti: **Barra** (la riapre),
**Schermo** (tutto schermo sì/no), **Navigazione** (mostra/nasconde Avanti, Indietro e Vai) e **Blocca**.

**Blocco dello schermo.** Con il lucchetto (disponibile con la Barra degli strumenti accesa) lo schermo ignora ogni tocco (le note continuano ad aggiornarsi).
Per **sbloccare tieni premuto 5 secondi** il pulsante rosso in alto: una barra si riempie e poi lo schermo si libera.
Se lasci prima, resta bloccato.

**Il telefono non va in standby** finché la pagina è aperta. Il primo tocco sulla pagina attiva la protezione;
se il telefono si spegne lo stesso, controlla di non avere il *risparmio energetico* attivo.

In basso: **Indietro** (arancione) e **Avanti** (verde) grandi, più il campo **Vai** per saltare a un numero di slide.
Se ci sono più presentazioni aperte, scegli quale seguire dal menu **Presentazione** (non disponibile con *Blocca schermo Remoto*).
Ogni telefono ricorda le sue scelte (stile del testo, cosa mostrare).

## Opzioni da riga di comando
`mefistofele.exe --port 8080 --pin 1234 --no-control --no-browser --demo`

- `--port` porta del server (predefinita 8080)
- `--pin` PIN per i comandi
- `--no-control` disattiva del tutto i comandi dal telefono
- `--no-browser` non apre la pagina di gestione sul PC
- `--demo` prova senza PowerPoint

## Se qualcosa non funziona
- **Il telefono non si collega**: stessa rete? Se il PC ha più schede, scegli quella giusta nella pagina di gestione. Firewall di Windows: consenti le *Reti private* (e pubbliche se il Wi-Fi è visto come pubblico).
- **Le note non compaiono**: la presentazione deve essere **in esecuzione** (a schermo intero), non solo aperta.
- **Le anteprime non compaiono**: apri `http://localhost:8080/diagnostica` sul PC, premi *Prova le anteprime ora* e *Copia tutto*.
- Gli errori sono scritti anche in `mefistofele.log`, accanto all'exe.

---
Ideato e sviluppato da [www.andreabattaglioli.it](https://www.andreabattaglioli.it) ·
Powered by [MyWebTvLive.com](https://www.mywebtvlive.com) · Made in Bolognina
