; ============================================================
;  Inno Setup - LAN Scanner [by Mauro Tassera]
; ============================================================
;  COME USARLO
;  1) Installa Inno Setup (gratuito): https://jrsoftware.org/isdl.php
;  2) Pubblica prima l'app in versione autonoma (comando unico, senza profilo):
;        dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:EnableCompressionInSingleFile=true
;     (crea  bin\Release\net8.0-windows\win-x64\publish\LanScanner.exe  senza richiedere .NET)
;  3) Apri questo file con Inno Setup e premi Compila (oppure:  ISCC.exe LanScanner.iss)
;  4) Il setup pronto esce in:  installer\Setup-LANScanner-1.0.0.exe
; ============================================================

#define MyAppName "LAN Scanner"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Mauro Tassera"
#define MyAppExeName "LanScanner.exe"
; Cartella con l'eseguibile pubblicato (relativa a questo .iss)
#define MyPublishDir "bin\Release\net8.0-windows\win-x64\publish"

[Setup]
AppId={{8C4B2F1A-5D3E-4A69-B7C2-9E0F1A2B3C4D}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=installer
OutputBaseFilename=Setup-LANScanner-{#MyAppVersion}
SetupIconFile=antenna.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "italian"; MessagesFile: "compiler:Languages\Italian.isl"

[Tasks]
Name: "desktopicon"; Description: "Crea un'icona sul &desktop"; GroupDescription: "Icone aggiuntive:"; Flags: unchecked

[Files]
Source: "{#MyPublishDir}\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
; Database produttori completo, se presente accanto all'exe pubblicato
Source: "{#MyPublishDir}\oui.txt"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Disinstalla {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Avvia {#MyAppName}"; Flags: nowait postinstall skipifsilent
