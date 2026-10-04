#!/usr/bin/env bash
# ============================================================
#  离线安装依赖（bash / Python 3.11）
#  从工程内 vendor/ 读取 .whl，全程不联网。
#  依赖齐全时不做任何安装（幂等）；逻辑统一由 scripts/deps_check.py 提供。
# ============================================================
set -euo pipefail
cd "$(dirname "$0")"

PY="${PYTHON:-python3.11}"

echo "[1/2] 检查 Python 版本 ..."
"$PY" -c "import sys; assert sys.version_info[:2]==(3,11), 'Need Python 3.11, got %d.%d' % sys.version_info[:2]"

echo "[2/2] 离线依赖自检 / 安装（只用 vendor/，不联网）..."
"$PY" "scripts/deps_check.py" --offline

echo
echo "[OK] 完成。验证："
echo "     $PY -c \"import pandas, numpy, openpyxl; print(pandas.__version__, numpy.__version__, openpyxl.__version__)\""
