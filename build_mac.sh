#!/usr/bin/env bash
# TaskManager Mac Build Script Runner
set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

# 1. Python 가상환경 탐색
if [ -d ".venv" ]; then
    PYTHON=".venv/bin/python"
elif [ -d "venv" ]; then
    PYTHON="venv/bin/python"
elif command -v python3 &>/dev/null; then
    PYTHON="python3"
else
    echo "[!] python3를 찾을 수 없습니다."
    exit 1
fi

echo "=================================================="
echo " TMProject macOS App Builder"
echo " Python: $($PYTHON --version 2>&1) ($PYTHON)"
echo "=================================================="

# 2. PyInstaller 설치 여부 확인
if ! $PYTHON -c "import PyInstaller" &>/dev/null; then
    echo "[*] PyInstaller가 설치되어 있지 않습니다. requirements.txt 설치를 진행합니다..."
    $PYTHON -m pip install -r requirements.txt
fi

# 3. build_mac.py 실행
$PYTHON build_mac.py "$@"
