param([string]$ProjectRoot = 'F:\курсовая Копейкин Андрей\Копейкин Андрей\Копейкин Андрей Последняя версия — копия')
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

$diag = Join-Path $ProjectRoot 'docs\diagrams'
$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('bw_diagrams_' + (Get-Date -Format 'yyyyMMddHHmmss'))
New-Item -ItemType Directory -Path $tmp -Force | Out-Null

$cm = New-Object System.Drawing.Imaging.ColorMatrix
$cm.Matrix00 = 0.299; $cm.Matrix01 = 0.299; $cm.Matrix02 = 0.299
$cm.Matrix10 = 0.587; $cm.Matrix11 = 0.587; $cm.Matrix12 = 0.587
$cm.Matrix20 = 0.114; $cm.Matrix21 = 0.114; $cm.Matrix22 = 0.114
$cm.Matrix33 = 1.0; $cm.Matrix44 = 1.0

$map = New-Object System.Collections.ArrayList
foreach ($f in Get-ChildItem -LiteralPath $diag -File -Filter '*.png') {
    $src = [System.Drawing.Bitmap]::FromFile($f.FullName)
    $w = $src.Width; $h = $src.Height
    $dst = New-Object System.Drawing.Bitmap($w, $h, [System.Drawing.Imaging.PixelFormat]::Format24bppRgb)
    $attrs = New-Object System.Drawing.Imaging.ImageAttributes
    $attrs.SetColorMatrix($cm)
    $g = [System.Drawing.Graphics]::FromImage($dst)
    $g.Clear([System.Drawing.Color]::White)
    $g.DrawImage($src, (New-Object System.Drawing.Rectangle(0, 0, $w, $h)), 0, 0, $w, $h, [System.Drawing.GraphicsUnit]::Pixel, $attrs)
    $g.Dispose()
    $out = Join-Path $tmp $f.Name
    $dst.Save($out, [System.Drawing.Imaging.ImageFormat]::Png)
    $dst.Dispose(); $src.Dispose()
    [void]$map.Add([pscustomobject]@{
        Name   = $f.Name
        Source = $f.FullName
        BW     = $out
        MD5    = (Get-FileHash -LiteralPath $f.FullName -Algorithm MD5).Hash
    })
    Write-Output ('  {0,-24} {1}x{2} -> ч/Б' -f $f.Name, $w, $h)
}
$map | Export-Csv -LiteralPath (Join-Path $tmp 'map.csv') -NoTypeInformation -Encoding UTF8
Write-Output ('Временная папка: ' + $tmp)
Write-Output ('Обработано: ' + $map.Count)
