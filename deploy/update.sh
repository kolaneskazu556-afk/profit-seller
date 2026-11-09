#!/usr/bin/env bash
# Обновление: sudo bash /opt/wbozon/deploy/update.sh
set -euo pipefail
cd /opt/wbozon
git pull --ff-only
./venv/bin/pip install -r requirements.txt
chown -R wbozon:wbozon /opt/wbozon
systemctl restart wbozon
sleep 2; curl -fsS http://127.0.0.1:8000/ping && echo " — сервис работает"
