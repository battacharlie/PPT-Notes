@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo === Creazione di DeckWhisperer.exe ===
py -m pip install -r requirements.txt
if errorlevel 1 goto errore
py ..\tools\buildinfo.py
if errorlevel 1 goto errore
py -m PyInstaller --onefile --console --noconfirm --clean --name DeckWhisperer --icon assets\deckwhisperer.ico --add-data "assets;assets" --add-data "..\docs;docs" --version-file version_info.txt --hidden-import win32com.client --hidden-import pythoncom deckwhisperer.py
if errorlevel 1 goto errore
copy /Y dist\DeckWhisperer.exe DeckWhisperer.exe >nul
echo.
echo FATTO: DeckWhisperer.exe e' in questa cartella.
pause
exit /b 0
:errore
echo.
echo ERRORE durante la creazione. Controlla i messaggi sopra.
pause
exit /b 1
