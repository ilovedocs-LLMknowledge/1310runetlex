#!/bin/bash
# Запуск программы «Протокол разногласий» двойным щелчком в Finder
cd "$(dirname "$0")"
for PY in /opt/homebrew/bin/python3 /usr/local/bin/python3 /usr/bin/python3 python3; do
  if command -v "$PY" >/dev/null 2>&1; then exec "$PY" protocol_app.py; fi
done
echo "Не найден Python 3. Установите его с python.org и запустите снова."
read -r -p "Нажмите Enter, чтобы закрыть окно..."
