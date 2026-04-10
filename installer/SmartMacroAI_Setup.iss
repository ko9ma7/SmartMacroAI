; ============================================================
;  SmartMacroAI — Inno Setup Installer Script
;  Created by Phạm Duy - Giải pháp tự động hóa thông minh.
; ============================================================

#define AppName        "SmartMacroAI"
#define AppVersion     "1.1.0"
#define AppPublisher   "Phạm Duy"
#define AppURL         "https://www.facebook.com/neull"
#define AppExeName     "SmartMacroAI.exe"
#define SourceDir      "..\publish\win-x64"
#define AssetsDir      "..\Assets"
#define InstallerDir   "."

[Setup]
; ── Unique GUID — regenerate for production ──────────────────
AppId={{F3A8C27E-9B4D-4E1F-A3B2-8D5C7E9F2A1B}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} v{#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppURL}
AppSupportURL={#AppURL}
AppUpdatesURL={#AppURL}
AppCopyright=Created by Phạm Duy - Giải pháp tự động hóa thông minh.

; ── Install dir ───────────────────────────────────────────────
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
AllowNoIcons=no
DisableProgramGroupPage=no
DisableDirPage=no

; ── Output ───────────────────────────────────────────────────
OutputDir=..\installer_out
OutputBaseFilename=SmartMacroAI_Setup_v{#AppVersion}

; ── Icons ────────────────────────────────────────────────────
SetupIconFile={#InstallerDir}\logo.ico
WizardImageFile={#InstallerDir}\wizard_side.bmp
WizardSmallImageFile={#InstallerDir}\wizard_top.bmp

; ── Compression ──────────────────────────────────────────────
Compression=lzma2/ultra64
SolidCompression=yes
InternalCompressLevel=ultra64

; ── Privileges ───────────────────────────────────────────────
PrivilegesRequired=admin
PrivilegesRequiredOverridesAllowed=dialog

; ── Appearance ───────────────────────────────────────────────
WizardStyle=modern
WizardSizePercent=120

; ── Minimum OS: Windows 10 ───────────────────────────────────
MinVersion=10.0.17763

; ── Language ─────────────────────────────────────────────────
ShowLanguageDialog=no

; ── Uninstall ────────────────────────────────────────────────
UninstallDisplayIcon={app}\{#AppExeName}
UninstallDisplayName={#AppName} v{#AppVersion}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
; Desktop shortcut
Name: "desktopicon";        Description: "Tạo biểu tượng trên Desktop";      GroupDescription: "Biểu tượng:"; Flags: checkedonce
; Quick Launch (Win7 compat) / Start menu
Name: "startmenuicon";     Description: "Tạo biểu tượng trong Start Menu";   GroupDescription: "Biểu tượng:"; Flags: checkedonce

; ── Optional: run at Windows startup ──────────────────────────
; Name: "startup"; Description: "Khởi động cùng Windows"; GroupDescription: "Khởi động:"; Flags: unchecked

[Dirs]
; Runtime folders needed by the app
Name: "{app}\Scripts"
Name: "{app}\templates"
Name: "{app}\tessdata"

[Files]
; ── Main app binaries ─────────────────────────────────────────
Source: "{#SourceDir}\*";     DestDir: "{app}";   Flags: ignoreversion recursesubdirs createallsubdirs; Excludes: "*.pdb"

; ── App assets ────────────────────────────────────────────────
Source: "{#AssetsDir}\logo.ico";      DestDir: "{app}\Assets"; Flags: ignoreversion
Source: "{#AssetsDir}\logo.png";      DestDir: "{app}\Assets"; Flags: ignoreversion
Source: "{#AssetsDir}\qr_bank.png";   DestDir: "{app}\Assets"; Flags: ignoreversion

; ── tessdata (OCR) — included only when the folder exists at build time ───────
; Source: "..\tessdata\*"; DestDir: "{app}\tessdata"; Flags: ignoreversion recursesubdirs createallsubdirs; Check: TessdataExists

[Icons]
; ── Start Menu group ─────────────────────────────────────────
Name: "{group}\{#AppName}";                             Filename: "{app}\{#AppExeName}"; IconFilename: "{app}\Assets\logo.ico"; Comment: "SmartMacroAI - Giải pháp tự động hóa thông minh"
Name: "{group}\Gỡ cài đặt {#AppName}";                  Filename: "{uninstallexe}"

; ── Desktop ──────────────────────────────────────────────────
Name: "{autodesktop}\{#AppName}";                       Filename: "{app}\{#AppExeName}"; IconFilename: "{app}\Assets\logo.ico"; Comment: "SmartMacroAI - Giải pháp tự động hóa thông minh"; Tasks: desktopicon

[Run]
; ── Install Playwright Chromium after main files are extracted ─
Filename: "powershell.exe";  Parameters: "-NoProfile -ExecutionPolicy Bypass -File ""{app}\playwright.ps1"" install chromium"; \
  WorkingDir: "{app}"; \
  StatusMsg: "Cài đặt trình duyệt Playwright (Chromium)... Vui lòng đợi ~1-2 phút."; \
  Flags: waituntilterminated runhidden; \
  Description: "Cài đặt Playwright Chromium (cho Web automation)"

; ── Launch app after install ──────────────────────────────────
Filename: "{app}\{#AppExeName}"; \
  Description: "Khởi chạy {#AppName}"; \
  Flags: nowait postinstall skipifsilent

[UninstallRun]
; Kill running instances before uninstall
Filename: "taskkill.exe"; Parameters: "/F /IM {#AppExeName}"; Flags: runhidden; RunOnceId: "KillApp"

[UninstallDelete]
; Remove user-generated runtime data on uninstall (optional — comment out to keep user data)
Type: filesandordirs; Name: "{app}\Scripts"
Type: filesandordirs; Name: "{app}\templates"

[Code]
// ─── Check if tessdata folder exists (to skip the Files entry if absent) ─────
function TessdataExists: Boolean;
begin
  Result := DirExists(ExpandConstant('{src}\..\tessdata'));
end;

// ─── Pre-flight: check .NET 8 Runtime ────────────────────────────────────────
function IsDotNet8Installed: Boolean;
var
  Key: String;
  Value: String;
begin
  Key := 'SOFTWARE\dotnet\Setup\InstalledVersions\x64\sdk';
  Result := RegQueryStringValue(HKLM, Key, '', Value);
  if not Result then
  begin
    // Check runtime path instead
    Result := DirExists(ExpandConstant('{pf}\dotnet\shared\Microsoft.WindowsDesktop.App\8.0.0')) or
              DirExists(ExpandConstant('{pf64}\dotnet\shared\Microsoft.WindowsDesktop.App\8.0.0'));
  end;
end;

// ─── Wizard page: warn user if needed, but do NOT block (self-contained) ─────
procedure InitializeWizard;
begin
  // Nothing blocking — publish is self-contained so .NET is bundled
end;

// ─── Add app to Windows "Add / Remove Programs" properly ─────────────────────
procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    RegWriteStringValue(HKLM,
      'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\' + '{#AppExeName}',
      '', ExpandConstant('{app}\{#AppExeName}'));
  end;
end;
