#!/bin/bash
# Hook di certbot (/etc/letsencrypt/renewal-hooks/deploy/): dopo il rinnovo del certificato di
# remote.mefistofele.com copia i file per il relay e lo riavvia. Non fa nulla per gli altri certificati.
case "$RENEWED_LINEAGE" in
  */remote.mefistofele.com)
    install -o mefrelay -g mefrelay -m 640 "$RENEWED_LINEAGE/fullchain.pem" /etc/mefistofele-relay/fullchain.pem
    install -o mefrelay -g mefrelay -m 600 "$RENEWED_LINEAGE/privkey.pem" /etc/mefistofele-relay/privkey.pem
    systemctl restart mefistofele-relay
    ;;
esac
