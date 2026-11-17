param(
    [string]$ProjectRoot = 'F:\курсовая Копейкин Андрей\Копейкин Андрей\Копейкин Андрей Последняя версия — копия',
    [string]$BwDir = 'C:\Users\Andrey\AppData\Local\Temp\bw_diagrams_20261116024255'
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression.FileSystem

$doc = Join-Path $ProjectRoot 'Копейкин _пояснительная_записка_курсового_проекта.docx'
$diag = Join-Path $ProjectRoot 'docs\diagrams'

# --- 1. карта: md5 цветной исходной диаграммы -> ч/Б файл с тем же именем ---
$bwByMd5 = @{}
Get-ChildItem -LiteralPath $diag -Filter '*.png' | ForEach-Object {
    $bwFile = Join-Path $BwDir $_.Name
    if (Test-Path -LiteralPath $bwFile) {
        $bwByMd5[(Get-FileHash -LiteralPath $_.FullName -Algorithm MD5).Hash] = $bwFile
    }
}
Write-Output ('Ч/Б файлов подготовлено: ' + $bwByMd5.Count)

# --- 2. grayscale-матрица ---
$cm = New-Object System.Drawing.Imaging.ColorMatrix
$cm.Matrix00 = 0.299; $cm.Matrix01 = 0.299; $cm.Matrix02 = 0.299
$cm.Matrix10 = 0.587; $cm.Matrix11 = 0.587; $cm.Matrix12 = 0.587
$cm.Matrix20 = 0.114; $cm.Matrix21 = 0.114; $cm.Matrix22 = 0.114
$cm.Matrix33 = 1.0; $cm.Matrix44 = 1.0

function Convert-ToGrayBytes([byte[]]$bytes) {
    $ms = New-Object System.IO.MemoryStream(,$bytes)
    $src = New-Object System.Drawing.Bitmap($ms)
    $w = $src.Width; $h = $src.Height
    $dst = New-Object System.Drawing.Bitmap($w, $h, [System.Drawing.Imaging.PixelFormat]::Format24bppRgb)
    $attrs = New-Object System.Drawing.Imaging.ImageAttributes
    $attrs.SetColorMatrix($cm)
    $g = [System.Drawing.Graphics]::FromImage($dst)
    $g.Clear([System.Drawing.Color]::White)
    $g.DrawImage($src, (New-Object System.Drawing.Rectangle(0, 0, $w, $h)), 0, 0, $w, $h, [System.Drawing.GraphicsUnit]::Pixel, $attrs)
    $g.Dispose()
    $out = New-Object System.IO.MemoryStream
    $dst.Save($out, [System.Drawing.Imaging.ImageFormat]::Png)
    $result = $out.ToArray()
    $out.Dispose(); $dst.Dispose(); $src.Dispose(); $ms.Dispose()
    return $result
}

# --- 3. замена media в docx ---
$zip = [System.IO.Compression.ZipFile]::Open($doc, 'Update')
try {
    $entries = @($zip.Entries | Where-Object { $_.FullName -like 'word/media/*.png' })
    foreach ($e in $entries) {
        $ms = New-Object System.IO.MemoryStream
        $st = $e.Open(); $st.CopyTo($ms); $st.Close()
        $bytes = $ms.ToArray(); $ms.Dispose()
        $sha = [System.Security.Cryptography.MD5]::Create()
        $h = ([BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-', '')

        $newBytes = $null
        $label = ''
        if ($bwByMd5.ContainsKey($h)) {
            $newBytes = [System.IO.File]::ReadAllBytes($bwByMd5[$h])
            $label = 'диаграмма -> ч/Б (исходник)'
        }
        elseif ($e.FullName -eq 'word/media/image2.png') {
            # use_case: рендер внутри документа, переводим в ч/Б на месте
            $newBytes = Convert-ToGrayBytes $bytes
            $label = 'use_case -> ч/Б (рендер в документе)'
        }
        if ($null -eq $newBytes) { continue }

        $st2 = $e.Open()
        $st2.SetLength(0)
        $st2.Write($newBytes, 0, $newBytes.Length)
        $st2.Close()
        Write-Output ('  {0,-22} {1,8} -> {2,8} байт  {3}' -f (Split-Path $e.FullName -Leaf), $bytes.Length, $newBytes.Length, $label)
    }
}
finally { $zip.Dispose() }
Write-Output 'DOCX обновлён'
