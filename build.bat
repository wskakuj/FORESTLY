@echo off
REM ==========================================
REM Forestly - lokalna budowa EXE (GUI web)
REM ==========================================
REM Wymaga: pip install pyinstaller
REM Wymaga: pip install -r requirements.txt

echo Budowanie Forestly.exe...
pyinstaller --noconfirm --onefile --windowed --icon "forestly.ico" --name "Forestly" --add-data "STR_TYT.docx;." --add-data "STR_TYT_TYLKO-ISL-2.docx;." --add-data "Skroty.docx;." --add-data "opis_og_szablon.docx;." --add-data "opis_og_szablon_mazowiecka.docx;." --add-data "opis_og_szablon_taksator.docx;." --add-data "gdos_obszary.json;." --add-data "BIAŁYNIN KRASÓWKA.xlsx;." --add-data "config;config" --add-data "pusty;pusty" --add-data "webapp;webapp" --collect-all customtkinter --collect-all CTkToolTip --collect-all pandas --collect-all numpy --collect-all fitz --collect-all openpyxl --collect-all docx --collect-all PIL --collect-all pypdf --collect-all pyodbc --collect-all webview --collect-all clr_loader --hidden-import win32com.client --hidden-import pythoncom --hidden-import pyautogui --collect-submodules app main.py

echo.
if exist "dist\Forestly.exe" (
    echo Gotowe! Plik: dist\Forestly.exe
) else (
    echo BLAD budowy Forestly.exe
)
pause
