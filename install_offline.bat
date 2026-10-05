@echo off
REM ============================================================
REM  离线安装依赖（Windows / Python 3.11）
REM  从工程内 vendor\ 读取 .whl，全程不联网。
REM  依赖齐全时不做任何安装（幂等）；逻辑统一由 scripts\deps_check.py 提供。
REM ============================================================
setlocal
cd /d "%~dp0"

echo [1/2] 检查 Python 版本 ...
python -c "import sys; assert sys.version_info[:2]==(3,11), 'Need Python 3.11'" || (
  echo [ERROR] 需要 Python 3.11，请先安装并确保 python 指向 3.11
  exit /b 1
)

echo [2/2] 离线依赖自检 / 安装（只用 vendor\，不联网）...
python "scripts\deps_check.py" --offline
if errorlevel 1 (
  echo [ERROR] 离线安装失败
  exit /b 1
)

echo.
echo [OK] 完成。验证：
echo      python -c "import pandas, numpy, openpyxl; print(pandas.__version__, numpy.__version__, openpyxl.__version__)"
endlocal
