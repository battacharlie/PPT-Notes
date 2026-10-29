@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo === Creazione di reMNote.exe ===
py -m pip install -r requirements.txt
if errorlevel 1 goto errore
py -m PyInstaller --onefile --console --noconfirm --clean --name reMNote --icon assets\remnote.ico --add-data "assets;assets" --version-file version_info.txt --hidden-import win32com.client --hidden-import pythoncom remnote.py
if errorlevel 1 goto errore
copy /Y dist\reMNote.exe reMNote.exe >nul
echo.
echo FATTO: reMNote.exe e' in questa cartella.
pause
exit /b 0
:errore
echo.
echo ERRORE durante la creazione. Controlla i messaggi sopra.
pause
exit /b 1
