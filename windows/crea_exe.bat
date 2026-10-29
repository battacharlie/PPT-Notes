@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo === Creazione di PPTNotes.exe ===
py -m pip install -r requirements.txt
if errorlevel 1 goto errore
py -m PyInstaller --onefile --console --noconfirm --clean --name PPTNotes --hidden-import win32com.client --hidden-import pythoncom ppt_notes.py
if errorlevel 1 goto errore
copy /Y dist\PPTNotes.exe PPTNotes.exe >nul
echo.
echo FATTO: PPTNotes.exe e' in questa cartella.
pause
exit /b 0
:errore
echo.
echo ERRORE durante la creazione. Controlla i messaggi sopra.
pause
exit /b 1
