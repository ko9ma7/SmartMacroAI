[Setup]
; --- Thông tin phần mềm ---
AppName=SmartMacroAI
AppVersion=1.0.0
AppPublisher=Phạm Duy
AppPublisherURL=https://github.com/TroniePh/SmartMacroAI
AppSupportURL=https://www.facebook.com/neull

; --- Nơi cài đặt mặc định trên máy khách hàng ---
DefaultDirName={autopf}\SmartMacroAI
DefaultGroupName=SmartMacroAI

; --- File xuất ra ---
OutputBaseFilename=Setup_SmartMacroAI_v1.0.0
OutputDir=userdocs\Desktop
Compression=lzma2/ultra64
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64

; 👉 BƯỚC QUAN TRỌNG 1: Gắn Icon cho chính file Setup.exe
; (Đảm bảo bạn có file logo.ico trong thư mục này)
SetupIconFile=E:\macro\SmartMacroAI\bin\Release\net8.0-windows\win-x64\publish\asset\logo.ico

[Tasks]
Name: "desktopicon"; Description: "Tao bieu tuong tren man hinh (Desktop)"; GroupDescription: "Them:"; Flags: unchecked

[Files]
; --- Lấy toàn bộ dữ liệu từ thư mục publish ---
; (Lệnh này sẽ lấy file .exe, copy luôn cả thư mục tessdata và thư mục asset đi theo)
Source: "E:\macro\SmartMacroAI\bin\Release\net8.0-windows\win-x64\publish\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; 👉 BƯỚC QUAN TRỌNG 2: Tạo Shortcut và gắn Icon cho nó
; Shortcut trong Start Menu
Name: "{group}\SmartMacroAI"; Filename: "{app}\SmartMacroAI.exe"; IconFilename: "{app}\asset\logo.ico"
; Trình gỡ cài đặt
Name: "{group}\Go Cai Dat SmartMacroAI"; Filename: "{uninstallexe}"
; Shortcut ngoài Desktop (Khi người dùng ghim cái này xuống Taskbar, nó cũng sẽ lấy logo này)
Name: "{autodesktop}\SmartMacroAI"; Filename: "{app}\SmartMacroAI.exe"; Tasks: desktopicon; IconFilename: "{app}\asset\logo.ico"

[Run]
; --- Chạy phần mềm ngay sau khi cài xong ---
Filename: "{app}\SmartMacroAI.exe"; Description: "Mo SmartMacroAI"; Flags: nowait postinstall skipifsilent