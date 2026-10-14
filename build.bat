@echo off
rem Builds dist\Email to PDF.exe - the single file you share. Needs Python 3.10+.
cd /d "%~dp0"
if not exist .venv python -m venv .venv || exit /b 1
.venv\Scripts\python -m pip install --quiet --disable-pip-version-check extract-msg pywin32 sv-ttk pyinstaller || exit /b 1
.venv\Scripts\python test_email_to_pdf.py || exit /b 1
.venv\Scripts\python -m PyInstaller --noconfirm --onefile --windowed --collect-data sv_ttk --name "Email to PDF" email_to_pdf.py || exit /b 1
echo.
echo Built: %~dp0dist\Email to PDF.exe
