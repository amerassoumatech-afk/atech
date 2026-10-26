@echo off
echo === Building ShopPOS_V5.exe ===
where py >nul 2>nul || (echo Python is not installed. Install Python 3.12 from https://www.python.org/downloads/ ^(tick "Add Python to PATH"^) and run this file again. & pause & exit /b 1)
py -m pip install --upgrade pip
py -m pip install pyinstaller pywebview openpyxl pywin32 cryptography google-api-python-client google-auth-oauthlib google-auth-httplib2
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist ShopPOS_V5.spec del ShopPOS_V5.spec
py -m PyInstaller --clean --noconfirm --onefile --windowed --name ShopPOS_V5 ^
  --add-data "index.html;." ^
  --collect-all googleapiclient --collect-all google_auth_oauthlib ^
  --hidden-import win32com.client --hidden-import openpyxl --hidden-import cryptography ^
  --hidden-import pos_server --hidden-import pos_core ^
  launcher.py
echo.
echo Done. Your program is: dist\ShopPOS_V5.exe
echo Next time you only need to install the update zip from Settings ^> Updates - no rebuild.
echo Your data is NOT inside the exe: %%APPDATA%%\MobilePOS_V5  (backups: Documents\MobilePOS_V5_Backups)
pause
