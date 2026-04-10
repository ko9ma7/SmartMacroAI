# ============================================================
#  Build-Installer.ps1  —  SmartMacroAI packaging script
#  Usage:  .\Build-Installer.ps1
#  Created by Phạm Duy - Giải pháp tự động hóa thông minh.
# ============================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$Root   = $PSScriptRoot
$Pub    = "$Root\publish\win-x64"
$ISCC   = 'C:\Program Files (x86)\Inno Setup 6\ISCC.exe'
$ISS    = "$Root\installer\SmartMacroAI_Setup.iss"
$OutDir = "$Root\installer_out"

Write-Host "`n[1/4]  Generating wizard bitmaps..." -ForegroundColor Cyan
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$Root\installer\MakeBitmaps.ps1"

Write-Host "`n[2/4]  Publishing .NET 8 self-contained release build..." -ForegroundColor Cyan
& dotnet publish "$Root\SmartMacroAI.csproj" `
    -c Release -r win-x64 --self-contained true `
    -p:PublishSingleFile=false `
    -o "$Pub"

Write-Host "`n[3/4]  Compiling Inno Setup installer..." -ForegroundColor Cyan
New-Item -ItemType Directory -Force $OutDir | Out-Null
& $ISCC $ISS

Write-Host "`n[4/4]  Installer ready:" -ForegroundColor Green
Get-ChildItem $OutDir -Filter '*.exe' | ForEach-Object {
    $mb = [int]($_.Length / 1MB)
    Write-Host ("   $($_.Name)  [$mb MB]") -ForegroundColor Yellow
}
