@echo off
REM ==========================================
REM Forestly - lokalna budowa EXE (oba GUI)
REM ==========================================
REM Wymaga: pip install pyinstaller
REM Wymaga: pip install -r requirements.txt

echo Budowanie Forestly.exe (nowe GUI web)...
pyinstaller --noconfirm --onefile --windowed --icon "forestly.ico" --name "Forestly" --add-data "STR_TYT.docx;." --add-data "STR_TYT_TYLKO-ISL-2.docx;." --add-data "Skroty.docx;." --add-data "opis_og_szablon.docx;." --add-data "gdos_obszary.json;." --add-data "BIAŁYNIN KRASÓWKA.xlsx;." --add-data "config;config" --add-data "pusty;pusty" --add-data "webapp;webapp" --collect-all customtkinter --collect-all CTkToolTip --collect-all pandas --collect-all numpy --collect-all fitz --collect-all openpyxl --collect-all docx --collect-all PIL --collect-all pypdf --collect-all pyodbc --collect-all webview --collect-all clr_loader --hidden-import win32com.client --hidden-import pythoncom --hidden-import pyautogui --collect-submodules app main_web.py

echo Budowanie Forestly_OLD.exe (klasyczne GUI)...
pyinstaller --noconfirm --onefile --windowed --icon "kombajn.ico" --name "Forestly_OLD" --add-data "STR_TYT.docx;." --add-data "STR_TYT_TYLKO-ISL-2.docx;." --add-data "Skroty.docx;." --add-data "BIAŁYNIN KRASÓWKA.xlsx;." --add-data "config;config" --add-data "pusty;pusty" --collect-all customtkinter --collect-all CTkToolTip --collect-all pandas --collect-all numpy --collect-all fitz --collect-all openpyxl --collect-all docx --collect-all PIL --collect-all pypdf --collect-all pyodbc --hidden-import win32com.client --hidden-import pythoncom --hidden-import pyautogui --collect-submodules app main.py

echo.
if exist "dist\Forestly.exe" (
    echo Gotowe! Plik: dist\Forestly.exe
) else (
    echo BLAD budowy Forestly.exe
)
if exist "dist\Forestly_OLD.exe" (
    echo Gotowe! Plik: dist\Forestly_OLD.exe
) else (
    echo BLAD budowy Forestly_OLD.exe
)
pause
