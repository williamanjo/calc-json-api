# Draws the logo (assets/logo.png, 512 px) and the extension icon (images/icon.png, 42 px).
Add-Type -AssemblyName System.Drawing
$root = Split-Path $PSScriptRoot -Parent

function RoundedRect([float]$x, [float]$y, [float]$w, [float]$h, [float]$r) {
    $p = New-Object System.Drawing.Drawing2D.GraphicsPath
    $d = $r * 2
    $p.AddArc($x, $y, $d, $d, 180, 90)
    $p.AddArc($x + $w - $d, $y, $d, $d, 270, 90)
    $p.AddArc($x + $w - $d, $y + $h - $d, $d, $d, 0, 90)
    $p.AddArc($x, $y + $h - $d, $d, $d, 90, 90)
    $p.CloseFigure()
    return $p
}

$size = 512
$bmp = New-Object System.Drawing.Bitmap $size, $size
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.SmoothingMode = 'AntiAlias'
$g.TextRenderingHint = 'AntiAliasGridFit'
$g.Clear([System.Drawing.Color]::Transparent)

# background: Calc-like green gradient
$bg = RoundedRect 16 16 480 480 96
$grad = New-Object System.Drawing.Drawing2D.LinearGradientBrush (
    (New-Object System.Drawing.Point 0, 16), (New-Object System.Drawing.Point 0, 496),
    [System.Drawing.Color]::FromArgb(255, 67, 176, 42), [System.Drawing.Color]::FromArgb(255, 24, 110, 50))
$g.FillPath($grad, $bg)

# table card
$white = [System.Drawing.Brushes]::White
$g.FillPath($white, (RoundedRect 84 110 300 300 22))
$header = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255, 200, 236, 190))
$g.FillPath($header, (RoundedRect 84 110 300 72 22))
$g.FillRectangle($header, 84, 150, 300, 32)
$line = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(255, 24, 110, 50)), 6
foreach ($y in 182, 258, 334) { $g.DrawLine($line, 84, $y, 384, $y) }
foreach ($x in 184, 284) { $g.DrawLine($line, $x, 110, $x, 410) }
$g.DrawPath($line, (RoundedRect 84 110 300 300 22))

# "{ }" badge
$badge = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255, 32, 48, 64))
$g.FillEllipse($badge, 286, 286, 190, 190)
$g.DrawEllipse((New-Object System.Drawing.Pen ([System.Drawing.Color]::White), 10), 286, 286, 190, 190)
$font = New-Object System.Drawing.Font 'Consolas', 104, ([System.Drawing.FontStyle]::Bold), ([System.Drawing.GraphicsUnit]::Pixel)
$fmt = New-Object System.Drawing.StringFormat
$fmt.Alignment = 'Center'; $fmt.LineAlignment = 'Center'
$g.DrawString('{}', $font, $white, (New-Object System.Drawing.RectangleF 286, 290, 190, 190), $fmt)
$g.Dispose()

New-Item -ItemType Directory -Force "$root\assets", "$root\images" | Out-Null
$bmp.Save("$root\assets\logo.png", [System.Drawing.Imaging.ImageFormat]::Png)

$icon = New-Object System.Drawing.Bitmap 42, 42
$gi = [System.Drawing.Graphics]::FromImage($icon)
$gi.InterpolationMode = 'HighQualityBicubic'
$gi.SmoothingMode = 'AntiAlias'
$gi.DrawImage($bmp, 0, 0, 42, 42)
$gi.Dispose()
$icon.Save("$root\images\icon.png", [System.Drawing.Imaging.ImageFormat]::Png)
"ok"
