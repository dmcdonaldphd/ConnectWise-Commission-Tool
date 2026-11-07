#ifndef AppVersion
  #define AppVersion "1.1.0"
#endif
#ifndef PackageDir
  #define PackageDir "..\dist\package"
#endif

[Setup]
; Keep this ID unchanged across updates.
AppId={{E2A3D376-B63E-4C84-9B61-90CE5FDC8D22}
AppName=PHDTech Commission Tool
AppVersion={#AppVersion}
AppPublisher=PHD Communications, LLC
DefaultDirName={localappdata}\Programs\PHDTech Commission Tool
DefaultGroupName=PHDTech Commission Tool
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0
WizardStyle=modern
Compression=lzma2/normal
SolidCompression=yes
OutputDir=..\dist
OutputBaseFilename=PHDTech-Commission-Tool-Setup-{#AppVersion}
UninstallDisplayIcon={app}\runtime\node.exe
SetupLogging=yes
RestartApplications=no

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; Flags: unchecked

[Files]
Source: "{#PackageDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\PHDTech Commission Tool"; Filename: "{app}\Run Commission Tool.cmd"; WorkingDir: "{app}"
Name: "{autodesktop}\PHDTech Commission Tool"; Filename: "{app}\Run Commission Tool.cmd"; WorkingDir: "{app}"; Tasks: desktopicon

[Run]
Filename: "{app}\Run Commission Tool.cmd"; WorkingDir: "{app}"; Description: "Launch PHDTech Commission Tool"; Flags: shellexec nowait postinstall skipifsilent
