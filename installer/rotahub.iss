; Inno Setup Script - RotaHub Desktop
; Ajuste AppVersion e rode no Inno Setup Compiler (ISCC).

#define MyAppName "RotaHub Desktop"
#define MyAppVersionBase "4.1.0"
#define MyAppVersion MyAppVersionBase
#define MyAppPublisher "RotaHub"
#define MyAppExeName "RotaHubDesktop.exe"

[Setup]
AppId={{1E93499B-0A30-4A63-A6B9-3B69CA5F40A1}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\RotaHub Desktop
DefaultGroupName=RotaHub Desktop
DisableProgramGroupPage=yes
OutputDir=..\dist_installer
OutputBaseFilename=RotaHubDesktop_Setup_{#MyAppVersion}
SetupIconFile=..\assets\app_icon.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
Compression=lzma
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Tasks]
Name: "desktopicon"; Description: "Criar atalho na area de trabalho"; GroupDescription: "Atalhos:"

[Files]
Source: "..\dist\RotaHubDesktop\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\RotaHub Desktop"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\RotaHub Desktop"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Executar RotaHub Desktop"; Flags: nowait postinstall skipifsilent
