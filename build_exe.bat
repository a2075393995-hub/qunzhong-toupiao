@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
  echo 未找到 Python。请先安装 Python 3.10 或更高版本。
  pause
  exit /b 1
)

python -m pip install -r requirements.txt -r requirements-dev.txt
if errorlevel 1 goto :failed

python -m unittest discover -s tests -v
if errorlevel 1 goto :failed

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0tools\build_portable.ps1"
if errorlevel 1 goto :failed

echo 已生成：%~dp0dist\QunzhongVote-v0.3.6-Windows-Portable.zip
pause
exit /b 0

:failed
echo 构建失败，请查看上方错误信息。
pause
exit /b 1
