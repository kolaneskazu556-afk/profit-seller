#!/usr/bin/env bash
# Установка на чистый Ubuntu 24.04 (REG.RU VPS). Идемпотентен: можно запускать повторно.
#   curl -fsSL https://raw.githubusercontent.com/<USER>/<REPO>/main/setup.sh -o setup.sh
#   sudo DOMAIN=example.ru REPO_URL=https://github.com/<USER>/<REPO>.git CERTBOT_EMAIL=you@mail.ru bash setup.sh
# Перед запуском A-запись домена должна указывать на IP сервера (иначе не выдастся HTTPS-сертификат).
# Ключ GigaChat вводится на сервере в консоли и в чат не пересылается.
set -euo pipefail
: "${DOMAIN:?Укажите DOMAIN=ваш-домен.ru}"
: "${REPO_URL:?Укажите REPO_URL=https://github.com/<USER>/<REPO>.git}"
EMAIL="${CERTBOT_EMAIL:-}"
APP=/opt/wbozon; DATA=/var/lib/wbozon; USR=wbozon
[ "$(id -u)" -eq 0 ] || { echo "Запустите через sudo"; exit 1; }
export DEBIAN_FRONTEND=noninteractive
echo "== 1/8 Пакеты"; apt-get update -y
apt-get install -y python3 python3-venv python3-pip nginx certbot python3-certbot-nginx git ufw fonts-dejavu-core sqlite3 gnupg rclone curl openssl
echo "== 2/8 Пользователь и каталоги"
id -u $USR >/dev/null 2>&1 || useradd --system --home $APP --shell /usr/sbin/nologin $USR
mkdir -p $DATA/backups $APP; chown -R $USR:$USR $DATA
echo "== 3/8 Код"
if [ -d $APP/.git ]; then git -C $APP pull --ff-only; else git clone "$REPO_URL" $APP; fi
echo "== 4/8 Python-окружение"
[ -d $APP/venv ] || python3 -m venv $APP/venv
$APP/venv/bin/pip install --upgrade pip
$APP/venv/bin/pip install -r $APP/requirements.txt
echo "== 5/8 Настройки (.env)"
mkdir -p $APP/certs
if [ ! -f $APP/certs/russian_ca.pem ]; then
  if curl -fsSL -m 30 https://gu-st.ru/content/Other/doc/russian_trusted_root_ca.cer -o /tmp/ru_ca.cer; then
    openssl x509 -inform DER -in /tmp/ru_ca.cer -out $APP/certs/russian_ca.pem 2>/dev/null || cp /tmp/ru_ca.cer $APP/certs/russian_ca.pem
  else echo "Не удалось скачать сертификат Минцифры: положите PEM в $APP/certs/russian_ca.pem вручную"; fi
fi
if [ ! -f $APP/.env ]; then
  GKEY=""
  if [ -t 0 ]; then read -rsp "Ключ авторизации GigaChat (Enter — пропустить, чат будет отключён): " GKEY; echo; fi
  cp $APP/.env.example $APP/.env
  sed -i "s|^GIGACHAT_AUTH_KEY=.*|GIGACHAT_AUTH_KEY=${GKEY}|" $APP/.env
fi
chown -R $USR:$USR $APP; chmod 600 $APP/.env
echo "== 6/8 Служба"
cat > /etc/systemd/system/wbozon.service <<UNIT
[Unit]
Description=Profit seller service
After=network.target
[Service]
User=$USR
WorkingDirectory=$APP
EnvironmentFile=$APP/.env
ExecStart=$APP/venv/bin/uvicorn app:app --host 127.0.0.1 --port 8000 --workers 1
Restart=always
NoNewPrivileges=true
ProtectSystem=full
ReadWritePaths=$DATA
[Install]
WantedBy=multi-user.target
UNIT
systemctl daemon-reload; systemctl enable --now wbozon; sleep 3
curl -fsS http://127.0.0.1:8000/ping && echo " — приложение запущено" || { echo "Приложение не отвечает: journalctl -u wbozon -n 50"; exit 1; }
echo "== 7/8 Nginx и HTTPS"
cat > /etc/nginx/sites-available/wbozon <<'NGX'
limit_req_zone $binary_remote_addr zone=wbozon:10m rate=10r/s;
server {
    listen 80;
    server_name DOMAIN_PLACEHOLDER;
    server_tokens off;
    client_max_body_size 25m;
    add_header X-Content-Type-Options nosniff always;
    add_header X-Frame-Options DENY always;
    add_header Referrer-Policy no-referrer always;
    location / {
        limit_req zone=wbozon burst=30 nodelay;
        proxy_pass http://127.0.0.1:8000;
        proxy_set_header Host $host;
        proxy_set_header X-Forwarded-For $remote_addr;
        proxy_set_header X-Forwarded-Proto $scheme;
        proxy_read_timeout 120s;
    }
}
NGX
sed -i "s/DOMAIN_PLACEHOLDER/$DOMAIN/" /etc/nginx/sites-available/wbozon
ln -sf /etc/nginx/sites-available/wbozon /etc/nginx/sites-enabled/wbozon; rm -f /etc/nginx/sites-enabled/default
nginx -t && systemctl reload nginx
if [ -n "$EMAIL" ]; then CB="-m $EMAIL"; else CB="--register-unsafely-without-email"; fi
certbot --nginx -d "$DOMAIN" --non-interactive --agree-tos $CB --redirect \
  || echo "!! HTTPS не выдан (проверьте, что DNS домена указывает на этот сервер). Повторить: certbot --nginx -d $DOMAIN"
ufw allow OpenSSH; ufw allow 'Nginx Full'; ufw --force enable
echo "== 8/8 Резервные копии"
if [ ! -f /root/.wbozon_backup_pass ]; then
  openssl rand -base64 32 > /root/.wbozon_backup_pass; chmod 600 /root/.wbozon_backup_pass
  echo "Пароль шифрования копий: файл /root/.wbozon_backup_pass. Сохраните его в менеджер паролей (без него копии не расшифровать). Никому не пересылайте."
fi
chmod +x $APP/deploy/*.sh
echo "15 3 * * * root /opt/wbozon/deploy/backup.sh >> /var/log/wbozon-backup.log 2>&1" > /etc/cron.d/wbozon-backup
echo
echo "Готово: https://$DOMAIN"
echo "Осталось вручную: 1) rclone config (создайте хранилище с именем 'offsite' в российском S3), 2) монитор на https://$DOMAIN/ping (UptimeRobot/Healthchecks), 3) проверка: curl https://$DOMAIN/health/llm"
