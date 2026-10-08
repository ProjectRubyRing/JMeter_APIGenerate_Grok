#!/usr/bin/env bash
# Generate a JMeter 5.6.3 test plan on RHEL 9.8.
# Requires bash 4.2+ and python3 (AppStream). No pip packages are required.
# CSV and Excel parsing, JMX generation, and the usage workbook are implemented
# with the Python 3 standard library. This shell is the supported entry point.

set -euo pipefail

export LANG="${LANG:-C.UTF-8}"
export LC_ALL="${LC_ALL:-C.UTF-8}"
export PYTHONUNBUFFERED=1

usage() {
  cat <<'EOF'
usage: generate-jmeter.sh --input FILE --output-dir DIR [options]
       generate-jmeter.sh --init-templates --output-dir DIR

Excel (.xlsx) または CSV から、Apache JMeter 5.6.3 の JMX と利用手順 Excel を生成します。
初期値では有効な各 API をスレッド 1、ループ 1 で 1 回呼び出します。

options:
  -i, --input FILE       入力 .xlsx または .csv
  -o, --output-dir DIR   出力ディレクトリ
      --run              生成後に JMeter を非GUI実行し、JTL、jmeter.log、HTMLレポートを書く
      --jmeter PATH      jmeter 実行ファイル。未指定時は JMETER_HOME と PATH
      --threads N        スレッド数（未指定時は入力値、さらに無ければ 1）
      --loops N          ループ回数（未指定時は入力値、さらに無ければ 1）
      --ramp-up N        ランプアップ秒
      --init-templates   入力テンプレートを --output-dir に書き出す
  -h, --help             このヘルプ

RHEL 9.8:
  sudo dnf install -y python3 java-17-openjdk-headless
  export JMETER_HOME=/opt/apache-jmeter-5.6.3
EOF
}

to_native() {
  local path="$1"
  if [[ "${PY_PLATFORM:-}" == "win32" ]] && command -v cygpath >/dev/null 2>&1; then
    cygpath -w "$path"
    return
  fi
  printf '%s\n' "$path"
}

if [[ "${1:-}" == "-h" || "${1:-}" == "--help" ]]; then
  usage
  exit 0
fi

SCRIPT_DIR=$(CDPATH= cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
ROOT=$(CDPATH= cd -- "$SCRIPT_DIR/.." && pwd)
PY_SCRIPT="$ROOT/lib/jmeter_gen/__main__.py"

if [[ ! -f "$ROOT/lib/jmeter_gen/__main__.py" ]]; then
  printf 'error: generator not found: %s\n' "$ROOT/lib/jmeter_gen/__main__.py" >&2
  exit 1
fi

if [[ -n "${PYTHON:-}" ]]; then
  PY="$PYTHON"
elif command -v python3 >/dev/null 2>&1; then
  PY=$(command -v python3)
elif command -v python >/dev/null 2>&1; then
  PY=$(command -v python)
else
  cat >&2 <<'EOF'
error: python3 が見つかりません。
RHEL 9.8 では次で導入できます。
  sudo dnf install -y python3
EOF
  exit 127
fi

if ! "$PY" -c 'import sys; raise SystemExit(0 if sys.version_info >= (3, 9) else 1)'; then
  printf 'error: Python 3.9 以上が必要です。RHEL 9 の python3 は 3.9 です。\n' >&2
  exit 127
fi

PY_PLATFORM=$("$PY" -c 'import sys; print(sys.platform)')
PY_SCRIPT=$(to_native "$ROOT/lib/jmeter_gen/__main__.py")
if [[ "$PY_PLATFORM" == "win32" ]]; then
  export PYTHONPATH="$(to_native "$ROOT/lib")"
else
  export PYTHONPATH="$ROOT/lib${PYTHONPATH:+:$PYTHONPATH}"
fi
exec "$PY" "$PY_SCRIPT" "$@"
