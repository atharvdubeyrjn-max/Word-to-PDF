; Inno Setup installer for Word → PDF
#define MyAppName "Word → PDF"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Word PDF Helper"
#define MyAppExeName "WordToPDF.exe"

[Setup]
AppId={{B9D5A4B8-6A1D-4D36-9B20-2F1F7A7A0D19}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\Word PDF Helper
DefaultGroupName={#MyAppName}
OutputDir=installer
OutputBaseFilename=WordToPDF-Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64
PrivilegesRequired=admin

[Files]
Source: "publish\WordToPDF.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Word → PDF"; Filename: "{app}\WordToPDF.exe"
Name: "{autodesktop}\Word → PDF"; Filename: "{app}\WordToPDF.exe"

[Run]
Filename: "{app}\WordToPDF.exe"; Description: "Launch Word → PDF"; Flags: nowait postinstall skipifsilent
