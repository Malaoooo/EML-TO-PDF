@echo off
rem Builds dist\Email to PDF.exe (the single file you share), an unpacked copy in dist\app\ (faster to start,
rem used by the installer) and, if Inno Setup is installed, dist\Email-to-PDF-Setup.exe. Needs Python 3.10+.
cd /d "%~dp0"
if not exist .venv python -m venv .venv || exit /b 1
.venv\Scripts\python -m pip install --quiet --disable-pip-version-check extract-msg pywin32 sv-ttk tkinterdnd2 pypdf pyinstaller || exit /b 1
.venv\Scripts\python test_email_to_pdf.py || exit /b 1
set PYI=.venv\Scripts\python -m PyInstaller --noconfirm --windowed --icon icon.ico --splash splash.png --add-data "icon.ico;." --add-data "icon.png;." --add-data "icons;icons" --collect-data sv_ttk --collect-all tkinterdnd2
%PYI% --onefile --workpath build\onefile --name "Email to PDF" email_to_pdf.py || exit /b 1
%PYI% --onedir --workpath build\onedir --distpath dist\app --name "Email to PDF" email_to_pdf.py || exit /b 1
echo.
echo Built: %~dp0dist\Email to PDF.exe
echo Built: %~dp0dist\app\Email to PDF\

set "ISCC=%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" (echo Inno Setup not found, skipping the installer. & exit /b 0)
for /f "tokens=2 delims='" %%v in ('findstr /b /c:"VERSION = " email_to_pdf.py') do set "VERSION=%%v"
"%ISCC%" /Q /DAppVersion=%VERSION% installer.iss || exit /b 1
echo Built: %~dp0dist\Email-to-PDF-Setup.exe
