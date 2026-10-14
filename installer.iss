; Inno Setup script: builds dist\Email-to-PDF-Setup.exe from the unpacked build in dist\app\Email to PDF\
; (unpacked starts much faster than the single-file exe, which matters on slow PCs).
; Per-user install, no admin rights needed. Build: ISCC /DAppVersion=1.2.3 installer.iss

#ifndef AppVersion
  #define AppVersion "0.0.0"
#endif

[Setup]
AppId={{6F0B8E2A-4C1D-4E5B-9A7E-3D2C1B0A9F81}
AppName=Email to PDF
AppVersion={#AppVersion}
AppPublisher=Malaoooo
AppPublisherURL=https://github.com/Malaoooo/EML-TO-PDF
DefaultDirName={localappdata}\Programs\Email to PDF
DefaultGroupName=Email to PDF
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=dist
OutputBaseFilename=Email-to-PDF-Setup
SetupIconFile=icon.ico
UninstallDisplayIcon={app}\Email to PDF.exe
WizardStyle=modern
Compression=lzma2
SolidCompression=yes

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Shortcuts:"

[Files]
Source: "dist\app\Email to PDF\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\Email to PDF"; Filename: "{app}\Email to PDF.exe"
Name: "{group}\Uninstall Email to PDF"; Filename: "{uninstallexe}"
Name: "{userdesktop}\Email to PDF"; Filename: "{app}\Email to PDF.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\Email to PDF.exe"; Description: "Open Email to PDF now"; Flags: nowait postinstall skipifsilent
