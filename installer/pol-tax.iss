; Pol-Tax installer, built by scripts/build.py:
;   ISCC /DMyAppVersion=1.2.3 /O<output dir> installer\pol-tax.iss
; Paths are relative to this file.

#ifndef MyAppVersion
  #error Pass the version: ISCC /DMyAppVersion=x.y.z (scripts/build.py does it)
#endif

#define MyAppName "Pol-Tax"
#define MyAppExeName "pol-tax.exe"

[Setup]
; Keep this AppId - it is how Windows recognizes installed Pol-Tax as the same app on upgrade
AppId={{8DEC85E8-B727-4D18-B95F-4ACE327BC2B9}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher=Pol-Tax
AppPublisherURL=https://github.com/GlockPL/pol-tax
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
OutputBaseFilename=Pol-Tax_{#MyAppVersion}
SetupIconFile=pol_tax_icon.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
Compression=lzma
SolidCompression=yes
WizardStyle=modern
Uninstallable=yes

[Languages]
Name: "polish"; MessagesFile: "compiler:Languages\Polish.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "..\dist\pol-tax\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[InstallDelete]
; Files of older builds that are no longer shipped
Type: files; Name: "{app}\resources\version.npy"
Type: files; Name: "{app}\resources\version.txt"

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent
