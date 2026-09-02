@echo off
setlocal
py -m pip install --upgrade pip
py -m pip install -r requirements.txt pyinstaller
pyinstaller --noconfirm --clean --onefile --windowed --name "论文格式检查助手" main.py
echo.
echo 构建完成：dist\论文格式检查助手.exe
pause

