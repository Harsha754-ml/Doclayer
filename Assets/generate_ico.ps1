Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName WindowsBase

$assetsDir = "m:\TEST\WordBarcodeStudio\Assets"
$icoPath = Join-Path $assetsDir "AppIcon.ico"
$pngPath = Join-Path $assetsDir "app_icon.png"

# Render a high-quality 256x256 bitmap using GDI+
$bmp256 = New-Object System.Drawing.Bitmap 256, 256, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
$g = [System.Drawing.Graphics]::FromImage($bmp256)
$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
$g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
$g.Clear([System.Drawing.Color]::Transparent)

# Draw stylized DocLayer icon on deep black rounded square background
$bgBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255, 15, 15, 18))
$borderPen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(255, 38, 38, 42)), 4

# Rounded rectangle background
$rect = New-Object System.Drawing.Rectangle 12, 12, 232, 232
$path = New-Object System.Drawing.Drawing2D.GraphicsPath
$radius = 48
$d = $radius * 2
$path.AddArc($rect.X, $rect.Y, $d, $d, 180, 90)
$path.AddArc($rect.Right - $d, $rect.Y, $d, $d, 270, 90)
$path.AddArc($rect.Right - $d, $rect.Bottom - $d, $d, $d, 0, 90)
$path.AddArc($rect.X, $rect.Bottom - $d, $d, $d, 90, 90)
$path.CloseFigure()

$g.FillPath($bgBrush, $path)
$g.DrawPath($borderPen, $path)

# Draw document sheets & barcode lines
# Document sheet 1 (White paper)
$sheetBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255, 255, 255, 255))
$sheetRect = New-Object System.Drawing.Rectangle 52, 48, 152, 160
$sheetPath = New-Object System.Drawing.Drawing2D.GraphicsPath
$sRadius = 16
$sd = $sRadius * 2
$sheetPath.AddArc($sheetRect.X, $sheetRect.Y, $sd, $sd, 180, 90)
$sheetPath.AddArc($sheetRect.Right - $sd, $sheetRect.Y, $sd, $sd, 270, 90)
$sheetPath.AddArc($sheetRect.Right - $sd, $sheetRect.Bottom - $sd, $sd, $sd, 0, 90)
$sheetPath.AddArc($sheetRect.X, $sheetRect.Bottom - $sd, $sd, $sd, 90, 90)
$sheetPath.CloseFigure()
$g.FillPath($sheetBrush, $sheetPath)

# Barcode graphic on paper
$barBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255, 18, 18, 20))
# QR-like corner squares
$g.FillRectangle($barBrush, 72, 68, 28, 28)
$g.FillRectangle($sheetBrush, 78, 74, 16, 16)
$g.FillRectangle($barBrush, 82, 78, 8, 8)

$g.FillRectangle($barBrush, 156, 68, 28, 28)
$g.FillRectangle($sheetBrush, 162, 74, 16, 16)
$g.FillRectangle($barBrush, 166, 78, 8, 8)

$g.FillRectangle($barBrush, 72, 116, 28, 28)
$g.FillRectangle($sheetBrush, 78, 122, 16, 16)
$g.FillRectangle($barBrush, 82, 126, 8, 8)

# Barcode lines & data tracks
$g.FillRectangle($barBrush, 110, 72, 36, 6)
$g.FillRectangle($barBrush, 110, 84, 24, 6)
$g.FillRectangle($barBrush, 140, 84, 6, 6)

# Horizontal lines
$g.FillRectangle($barBrush, 110, 116, 74, 6)
$g.FillRectangle($barBrush, 110, 128, 54, 6)
$g.FillRectangle($barBrush, 110, 140, 68, 6)

# Bottom barcode stripes (1D barcode band)
$g.FillRectangle($barBrush, 72, 162, 6, 26)
$g.FillRectangle($barBrush, 82, 162, 4, 26)
$g.FillRectangle($barBrush, 90, 162, 10, 26)
$g.FillRectangle($barBrush, 104, 162, 4, 26)
$g.FillRectangle($barBrush, 112, 162, 8, 26)
$g.FillRectangle($barBrush, 124, 162, 4, 26)
$g.FillRectangle($barBrush, 132, 162, 12, 26)
$g.FillRectangle($barBrush, 148, 162, 4, 26)
$g.FillRectangle($barBrush, 156, 162, 8, 26)
$g.FillRectangle($barBrush, 168, 162, 6, 26)
$g.FillRectangle($barBrush, 178, 162, 6, 26)

$g.Flush()
$bmp256.Save($pngPath, [System.Drawing.Imaging.ImageFormat]::Png)

# Generate multi-size ICO file (256, 128, 64, 48, 32, 16)
$sizes = @(256, 128, 64, 48, 32, 16)
$pngStreams = @()

foreach ($sz in $sizes) {
    $resized = New-Object System.Drawing.Bitmap $sz, $sz, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $rg = [System.Drawing.Graphics]::FromImage($resized)
    $rg.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $rg.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $rg.DrawImage($bmp256, 0, 0, $sz, $sz)
    $rg.Flush()
    
    $ms = New-Object System.IO.MemoryStream
    $resized.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    $pngStreams += ,@($sz, $ms.ToArray())
    $rg.Dispose()
    $resized.Dispose()
}

$fs = New-Object System.IO.FileStream $icoPath, ([System.IO.FileMode]::Create)
$bw = New-Object System.IO.BinaryWriter $fs

# ICONDIR Header
$bw.Write([uint16]0) # Reserved
$bw.Write([uint16]1) # Type 1 = ICO
$bw.Write([uint16]$sizes.Count) # Count of images

$offset = 6 + ($sizes.Count * 16)

# Write ICONDIRENTRY for each size
foreach ($item in $pngStreams) {
    $sz = $item[0]
    $bytes = $item[1]
    
    $bw.Write([byte]$(if ($sz -ge 256) { 0 } else { $sz })) # Width
    $bw.Write([byte]$(if ($sz -ge 256) { 0 } else { $sz })) # Height
    $bw.Write([byte]0) # Colors
    $bw.Write([byte]0) # Reserved
    $bw.Write([uint16]1) # Color planes
    $bw.Write([uint16]32) # Bit count (32bpp RGBA)
    $bw.Write([uint32]$bytes.Length) # Image bytes length
    $bw.Write([uint32]$offset) # Offset
    $offset += $bytes.Length
}

# Write PNG byte payloads
foreach ($item in $pngStreams) {
    $bytes = $item[1]
    $bw.Write($bytes)
}

$bw.Flush()
$bw.Close()
$fs.Close()
$g.Dispose()
$bmp256.Dispose()

Write-Host "Created AppIcon.ico and app_icon.png in $assetsDir" -ForegroundColor Green
