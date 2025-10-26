#!/usr/bin/env bash
set -euo pipefail

# -----------------------------------------------------------------------------
# tvoy_broker_convert.sh — запускает broker_to_target.py в изолированном venv и передает нужные флаги.
# Использует .venv в корне проекта, автоматически создаёт и наполняет его.
#
# Использование:
#   ./tvoy_broker_convert.sh <input.xlsx> [доп. аргументы]
#
# Примеры:
#   ./tvoy_broker_convert.sh report.xlsx
#   ./tvoy_broker_convert.sh report.xlsx --alloc-commission --sheet "Отчет"
#
# По умолчанию в Python-скрипт передаются:
#   --output tvoy_result.xlsx --debug --include-cash
# Доп. аргументы после имени файла пробрасываются как есть.
# -----------------------------------------------------------------------------

usage() {
  echo "Usage: $0 <input.xlsx> [extra args passed to broker_to_target.py]"
  exit 1
}

# 1) Проверим аргументы
if [[ $# -lt 1 ]]; then
  usage
fi

INPUT="$1"; shift
if [[ ! -f "$INPUT" ]]; then
  echo "Input file not found: $INPUT"
  exit 2
fi

# Абсолютные пути к входному файлу и его директории
# Работает и для относительных путей.
INPUT_DIR="$(cd "$(dirname "$INPUT")" && pwd)"
INPUT_BASENAME="$(basename "$INPUT")"
INPUT_ABS="$INPUT_DIR/$INPUT_BASENAME"

# Выходной файл в той же директории, имя — tvoy_result.xlsx
OUT_FILE="$INPUT_DIR/tvoy_result.xlsx"

# 2) Определим пути
SCRIPT_DIR="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" &>/dev/null && pwd)"
PROJECT_ROOT="$SCRIPT_DIR"
VENV_DIR="$PROJECT_ROOT/.venv"
PY_BIN="python3"
PIP_BIN="$VENV_DIR/bin/pip"
PYTHON="$VENV_DIR/bin/python"
REQ_FILE="$PROJECT_ROOT/requirements.txt"

# 3) Проверим, что python3 доступен
if ! command -v "$PY_BIN" >/dev/null 2>&1; then
  echo "python3 не найден в PATH. Установите Python 3.11/3.12."
  exit 3
fi

# 4) Создадим venv при необходимости
if [[ ! -d "$VENV_DIR" ]]; then
  echo "Создаю виртуальное окружение в $VENV_DIR ..."
  "$PY_BIN" -m venv "$VENV_DIR"
fi

# 5) Обновим pip и установим зависимости только если изменился requirements.txt
STAMP_FILE="$VENV_DIR/.req.hash"
if [[ -f "$REQ_FILE" ]]; then
  if command -v sha1sum >/dev/null 2>&1; then
    REQ_HASH="$(sha1sum "$REQ_FILE" | awk '{print $1}')"
  else
    # macOS: shasum по умолчанию
    REQ_HASH="$(shasum "$REQ_FILE" | awk '{print $1}')"
  fi
else
  REQ_HASH="no-req"
fi

NEED_INSTALL=0
if [[ ! -f "$STAMP_FILE" ]] || [[ "$(cat "$STAMP_FILE")" != "$REQ_HASH" ]]; then
  NEED_INSTALL=1
fi

if [[ "$NEED_INSTALL" -eq 1 ]]; then
  echo "Устанавливаю/обновляю зависимости в venv ..."
  "$PIP_BIN" install --upgrade pip
  if [[ -f "$REQ_FILE" ]]; then
    "$PIP_BIN" install -r "$REQ_FILE"
  else
    # Фолбэк: если нет requirements.txt — ставим минимальный набор
    "$PIP_BIN" install pandas==2.2.3 openpyxl==3.1.5
  fi
  echo "$REQ_HASH" > "$STAMP_FILE"
fi

# 6) Запуск целевого скрипта с фиксированными флагами и возможными доп. параметрами
echo "Запускаю broker_to_target.py ..."
"$PYTHON" "$PROJECT_ROOT/broker_to_target.py" \
  --input "$INPUT" \
  --output "$OUT_FILE" \
  --debug \
  --include-cash \
  "$@"

echo "Готово: $OUT_FILE"