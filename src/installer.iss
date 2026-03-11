; installer.iss — Inno Setup script for palmer-type.
;
; Packages the PyInstaller --onedir output (with bundled tectonic) into a
; Windows installer.  The version is injected at compile time via:
;
;   ISCC /DAppVersion=1.0.0 installer.iss
;
; Prerequisites:
;   - PyInstaller --onedir build must have been run first
;     (output in dist\palmer-type\)
;   - Inno Setup 6 (https://jrsoftware.org/isinfo.php)

#ifndef AppVersion
  #define AppVersion "dev"
#endif

[Setup]
AppName=palmer-type
AppVersion={#AppVersion}
AppVerName=palmer-type {#AppVersion}
AppPublisher=Yosuke Yamazaki
AppPublisherURL=https://github.com/yosukey/palmer-type
DefaultDirName={autopf}\palmer-type
DefaultGroupName=palmer-type
UninstallDisplayIcon={app}\palmer-type.exe
OutputDir=..\dist
OutputBaseFilename=palmer-type-{#AppVersion}-win-x64-setup
Compression=lzma2/ultra64
SolidCompression=yes
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
SetupIconFile=assets\palmer-type-installer.ico
WizardStyle=modern

[Files]
Source: "..\dist\palmer-type\*"; DestDir: "{app}"; Flags: recursesubdirs ignoreversion

[Icons]
Name: "{group}\palmer-type"; Filename: "{app}\palmer-type.exe"
Name: "{group}\Uninstall palmer-type"; Filename: "{uninstallexe}"
Name: "{autodesktop}\palmer-type"; Filename: "{app}\palmer-type.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional icons:"; Flags: unchecked

[Run]
Filename: "{app}\palmer-type.exe"; Description: "Launch palmer-type"; Flags: nowait postinstall skipifsilent

[Code]
// ---------------------------------------------------------------------------
// Uninstall: optionally remove Tectonic cache and user configuration
// ---------------------------------------------------------------------------

var
  ShouldCleanup: Boolean;

function InitializeUninstall(): Boolean;
begin
  Result := True;
  ShouldCleanup := (MsgBox(
    'Do you also want to remove the following user data?' + #13#10 + #13#10 +
    '  - Tectonic TeX cache  (%LOCALAPPDATA%\TectonicProject\Tectonic)' + #13#10 +
    '  - Application settings and font favorites  (%APPDATA%\palmer-type)' + #13#10 + #13#10 +
    'Note: debug log files (%APPDATA%\palmer-type\logs) will always be removed.' + #13#10 + #13#10 +
    'Click Yes to delete this data, or No to keep it.',
    mbConfirmation, MB_YESNO or MB_DEFBUTTON2) = IDYES);
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  TectonicDir: String;
  AppDataDir:  String;
  LogDir:      String;
begin
  if CurUninstallStep = usPostUninstall then
  begin
    // Always remove debug log files — no confirmation required.
    LogDir := ExpandConstant('{userappdata}\palmer-type\logs');
    if DirExists(LogDir) then
      DelTree(LogDir, True, True, True);

    if ShouldCleanup then
    begin
      // Remove Tectonic cache: %LOCALAPPDATA%\TectonicProject\Tectonic
      TectonicDir := ExpandConstant('{localappdata}\TectonicProject\Tectonic');
      if DirExists(TectonicDir) then
        DelTree(TectonicDir, True, True, True);
      // Also remove the parent if empty
      RemoveDir(ExpandConstant('{localappdata}\TectonicProject'));

      // Remove app config dir: %APPDATA%\palmer-type
      AppDataDir := ExpandConstant('{userappdata}\palmer-type');
      if DirExists(AppDataDir) then
        DelTree(AppDataDir, True, True, True);
    end;
  end;
end;
