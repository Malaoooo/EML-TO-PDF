@echo off
rem Builds dist\Email to PDF.exe (the single file you share) and, if Inno Setup is installed,
rem dist\Email-to-PDF-Setup.exe. Needs Python 3.10+.
cd /d "%~dp0"
if not exist .venv python -m venv .venv || exit /b 1
.venv\Scripts\python -m pip install --quiet --disable-pip-version-check extract-msg pywin32 sv-ttk tkinterdnd2 pypdf pyinstaller || exit /b 1
.venv\Scripts\python test_email_to_pdf.py || exit /b 1
.venv\Scripts\python -m PyInstaller --noconfirm --onefile --windowed --icon icon.ico --add-data "icon.ico;." --add-data "icon.png;." ^
    --collect-data sv_ttk --collect-all tkinterdnd2 --name "Email to PDF" email_to_pdf.py || exit /b 1
echo.
echo Built: %~dp0dist\Email to PDF.exe

set "ISCC=%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" (echo Inno Setup not found, skipping the installer. & exit /b 0)
for /f "tokens=2 delims='" %%v in ('findstr /b /c:"VERSION = " email_to_pdf.py') do set "VERSION=%%v"
"%ISCC%" /Q /DAppVersion=%VERSION% installer.iss || exit /b 1
echo Built: %~dp0dist\Email-to-PDF-Setup.exe
