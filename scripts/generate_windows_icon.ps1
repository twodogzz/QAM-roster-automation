param(
    [Parameter(Mandatory = $true)]
    [string]$SourcePng,

    [Parameter(Mandatory = $true)]
    [string]$DestinationIco
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Drawing

if (-not (Test-Path -LiteralPath $SourcePng)) {
    throw "Source PNG not found: $SourcePng"
}

$destinationDir = Split-Path -Parent $DestinationIco
if ($destinationDir -and -not (Test-Path -LiteralPath $destinationDir)) {
    New-Item -ItemType Directory -Path $destinationDir | Out-Null
}

$tempPng = [System.IO.Path]::Combine(
    [System.IO.Path]::GetTempPath(),
    ([System.IO.Path]::GetRandomFileName() + '.png')
)

$sourceImage = $null
$bitmap = $null
$graphics = $null

try {
    $sourceImage = [System.Drawing.Image]::FromFile($SourcePng)
    $bitmap = New-Object System.Drawing.Bitmap 256, 256
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.Clear([System.Drawing.Color]::Transparent)
    $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality

    $scale = [Math]::Min(256 / $sourceImage.Width, 256 / $sourceImage.Height)
    $drawWidth = [int][Math]::Round($sourceImage.Width * $scale)
    $drawHeight = [int][Math]::Round($sourceImage.Height * $scale)
    $drawX = [int][Math]::Floor((256 - $drawWidth) / 2)
    $drawY = [int][Math]::Floor((256 - $drawHeight) / 2)

    $graphics.DrawImage($sourceImage, $drawX, $drawY, $drawWidth, $drawHeight)
    $bitmap.Save($tempPng, [System.Drawing.Imaging.ImageFormat]::Png)
} finally {
    if ($graphics) { $graphics.Dispose() }
    if ($bitmap) { $bitmap.Dispose() }
    if ($sourceImage) { $sourceImage.Dispose() }
}

$pngBytes = [System.IO.File]::ReadAllBytes($tempPng)

$fileStream = $null
$writer = $null

try {
    $fileStream = [System.IO.File]::Open($DestinationIco, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write)
    $writer = New-Object System.IO.BinaryWriter($fileStream)

    # ICO header
    $writer.Write([UInt16]0)      # reserved
    $writer.Write([UInt16]1)      # type = icon
    $writer.Write([UInt16]1)      # image count

    # ICONDIRENTRY for a 256x256 PNG payload (0 means 256 in ICO format)
    $writer.Write([byte]0)        # width
    $writer.Write([byte]0)        # height
    $writer.Write([byte]0)        # color count
    $writer.Write([byte]0)        # reserved
    $writer.Write([UInt16]1)      # color planes
    $writer.Write([UInt16]32)     # bits per pixel
    $writer.Write([UInt32]$pngBytes.Length)
    $writer.Write([UInt32]22)     # header (6) + directory entry (16)

    $writer.Write($pngBytes)
} finally {
    if ($writer) { $writer.Dispose() }
    elseif ($fileStream) { $fileStream.Dispose() }

    if (Test-Path -LiteralPath $tempPng) {
        Remove-Item -LiteralPath $tempPng -Force
    }
}

Write-Host "Created icon: $DestinationIco"
