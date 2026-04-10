Add-Type -AssemblyName System.Drawing

$logoPath = 'E:\macro\SmartMacroAI\Assets\logo.png'

# ── Wizard side image: 164 x 314 px ──────────────────────────
$side  = New-Object System.Drawing.Bitmap(164, 314)
$g     = [System.Drawing.Graphics]::FromImage($side)
$rectAll = New-Object System.Drawing.Rectangle(0, 0, 164, 314)
$ptTop   = New-Object System.Drawing.Point(0, 0)
$ptBot   = New-Object System.Drawing.Point(164, 314)
$grad  = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
            $ptTop, $ptBot,
            [System.Drawing.Color]::FromArgb(30, 30, 46),
            [System.Drawing.Color]::FromArgb(49, 50, 68))
$g.FillRectangle($grad, $rectAll)

try {
    $logo    = [System.Drawing.Image]::FromFile($logoPath)
    $logoSz  = 80
    $lx      = [int]((164 - $logoSz) / 2)
    $dstRect = New-Object System.Drawing.Rectangle($lx, 32, $logoSz, $logoSz)
    $g.DrawImage($logo, $dstRect)
    $logo.Dispose()
} catch { Write-Warning "Could not load logo: $_" }

$font      = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
$brush     = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(205, 214, 244))
$fontSm    = New-Object System.Drawing.Font('Segoe UI', 7)
$brushSm   = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(166, 173, 200))
$sf        = New-Object System.Drawing.StringFormat
$sf.Alignment = [System.Drawing.StringAlignment]::Center

$g.DrawString('SmartMacroAI', $font,   $brush,   82, 126, $sf)
$g.DrawString('v1.1.0',       $fontSm, $brushSm, 82, 148, $sf)
$g.DrawString('by Pham Duy',  $fontSm, $brushSm, 82, 290, $sf)

$g.Dispose()
$side.Save('E:\macro\SmartMacroAI\installer\wizard_side.bmp',
           [System.Drawing.Imaging.ImageFormat]::Bmp)
$side.Dispose()

# ── Wizard header / small image: 55 x 55 px ──────────────────
$top   = New-Object System.Drawing.Bitmap(55, 55)
$g2    = [System.Drawing.Graphics]::FromImage($top)
$bgColor = [System.Drawing.Color]::FromArgb(30, 30, 46)
$g2.Clear($bgColor)

try {
    $logo2 = [System.Drawing.Image]::FromFile($logoPath)
    $dst2  = New-Object System.Drawing.Rectangle(4, 4, 47, 47)
    $g2.DrawImage($logo2, $dst2)
    $logo2.Dispose()
} catch { Write-Warning "Could not load logo2: $_" }

$g2.Dispose()
$top.Save('E:\macro\SmartMacroAI\installer\wizard_top.bmp',
          [System.Drawing.Imaging.ImageFormat]::Bmp)
$top.Dispose()

Write-Host 'Wizard bitmaps created successfully.'
