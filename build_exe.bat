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

python -m PyInstaller --noconfirm --clean VoteDocxApp.spec
if errorlevel 1 goto :failed

for /f %%v in ('python -c "from update_service import APP_VERSION; print(APP_VERSION)"') do set "APP_VERSION=%%v"
copy /Y "dist\QunzhongVote.exe" "群众投票_v%APP_VERSION%.exe" >nul
echo 已生成：%~dp0群众投票_v%APP_VERSION%.exe
pause
exit /b 0

:failed
echo 构建失败，请查看上方错误信息。
pause
exit /b 1
