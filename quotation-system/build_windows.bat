@echo off
chcp 65001 >nul
where pyinstaller >nul 2>nul
if errorlevel 1 (
  echo 请先在Windows电脑安装 Python 和 PyInstaller：
  echo python -m pip install pyinstaller pillow
  pause
  exit /b 1
)
if not exist dist mkdir dist
pyinstaller --noconfirm --clean --onefile --windowed --name 常青文创设计报价系统 --add-data "LOGO.png;." quotation_app.py
copy /Y "报价设置.json" "dist\报价设置.json" >nul 2>nul
echo 已生成 dist\常青文创设计报价系统.exe
pause
