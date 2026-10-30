# Relay per l'accesso remoto (server)

Il relay (`relay/mefistofele_relay.py`) è un piccolo programma in Python (solo libreria standard, Python 3.9 o successivo)
che mette in contatto il PC del relatore con i telefoni **fuori dalla rete**. Il PC si collega in uscita; i telefoni aprono
`https://remote.mefistofele.com:8443/s/<codice>/`. Inoltra **solo** pagina, stato, anteprime e comandi (elenco chiuso di percorsi)
e non salva nulla su disco.

## Installazione sul server (Rocky Linux 9, come utente root)
Prerequisiti: record DNS `remote.mefistofele.com` verso il server e certificato Let's Encrypt già ottenuto con
`certbot certonly --standalone -d remote.mefistofele.com`.

```
useradd --system --no-create-home --shell /sbin/nologin mefrelay
mkdir -p /opt/mefistofele-relay /etc/mefistofele-relay
curl -fsSL https://raw.githubusercontent.com/battacharlie/PPT-Notes/main/relay/mefistofele_relay.py -o /opt/mefistofele-relay/mefistofele_relay.py
curl -fsSL https://raw.githubusercontent.com/battacharlie/PPT-Notes/main/relay/mefistofele-relay.service -o /etc/systemd/system/mefistofele-relay.service
curl -fsSL https://raw.githubusercontent.com/battacharlie/PPT-Notes/main/relay/rinnovo-certificato.sh -o /etc/letsencrypt/renewal-hooks/deploy/mefistofele-relay.sh
chmod 755 /etc/letsencrypt/renewal-hooks/deploy/mefistofele-relay.sh
RENEWED_LINEAGE=/etc/letsencrypt/live/remote.mefistofele.com /etc/letsencrypt/renewal-hooks/deploy/mefistofele-relay.sh || true
firewall-cmd --permanent --add-port=8443/tcp && firewall-cmd --reload
systemctl daemon-reload && systemctl enable --now mefistofele-relay
```

Controlli: `systemctl status mefistofele-relay` e `curl -s https://remote.mefistofele.com:8443/health`.
Registro: `journalctl -u mefistofele-relay -f`.

## Sicurezza in breve
- Il link contiene un codice casuale di 24 caratteri generato dal PC a ogni attivazione; il PC si autentica con una chiave
  segreta (sul server resta solo l'impronta SHA-256): nessuno può prendere il suo posto.
- Da fuori i comandi richiedono sempre il PIN; dopo 5 PIN sbagliati si bloccano per 5 minuti.
- Limiti: 300 sessioni, 60 telefoni per sessione, limite di richieste per indirizzo, sessioni chiuse dopo 5 minuti senza il PC.
- Il servizio gira con un utente senza privilegi, con limiti di memoria e di CPU.
