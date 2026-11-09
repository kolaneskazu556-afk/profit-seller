#!/usr/bin/env bash
# Ежедневная зашифрованная копия базы. Запускается cron от root (ставится setup.sh).
set -euo pipefail
DATA=/var/lib/wbozon; OUT=$DATA/backups; PASS=/root/.wbozon_backup_pass
TS=$(date +%Y%m%d-%H%M%S)
mkdir -p "$OUT"
sqlite3 "$DATA/data.db" ".backup '$OUT/data-$TS.db'"
gpg --batch --yes --pinentry-mode loopback --passphrase-file "$PASS" -c --cipher-algo AES256 -o "$OUT/data-$TS.db.gpg" "$OUT/data-$TS.db"
rm -f "$OUT/data-$TS.db"
find "$OUT" -name 'data-*.db.gpg' -mtime +7 -delete
HC_URL=$(grep -E '^HC_URL=' /opt/wbozon/.env 2>/dev/null | cut -d= -f2- || true)
if rclone listremotes | grep -q '^offsite:'; then
  rclone copy "$OUT/data-$TS.db.gpg" offsite:wbozon-backups/
  rclone delete offsite:wbozon-backups/ --min-age 7d
  [ -n "$HC_URL" ] && curl -fsS -m 10 "$HC_URL" >/dev/null || true
else
  echo "rclone-хранилище 'offsite' не настроено: копия есть только на этом сервере" >&2
fi
