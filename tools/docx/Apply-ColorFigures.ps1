$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Drawing

$root = 'F:\курсовая Копейкин Андрей\Копейкин Андрей\Копейкин Андрей Последняя версия — копия'
$work = Join-Path $root 'Копейкин _пояснительная_записка_курсового_проекта.docx'
$diag = Join-Path $root 'docs\diagrams'
$stamp = Get-Date -Format 'yyyy-MM-dd-HHmmss'
$backup = Join-Path $root ("docs\backup_zapiska_before_color_$stamp.docx")
Copy-Item -LiteralPath $work -Destination $backup -Force
Write-Output "Бэкап: $backup"

$tmp = Join-Path $env:TEMP ("docx_colorfix_$stamp")
[System.IO.Compression.ZipFile]::ExtractToDirectory($work, $tmp)

# media -> цветной исходник (те же диаграммы, что в презентации)
$map = @{
    'word\media\image7.png'  = (Join-Path $diag 'architecture.png')
    'word\media\image20.png' = (Join-Path $diag 'alg_create.png')
    'word\media\image24.png' = (Join-Path $diag 'er.png')
}

foreach ($k in $map.Keys) {
    $dst = Join-Path $tmp $k
    $src = $map[$k]
    if (-not (Test-Path -LiteralPath $dst)) { throw "нет $dst" }
    if (-not (Test-Path -LiteralPath $src)) { throw "нет $src" }
    Copy-Item -LiteralPath $src -Destination $dst -Force
    $b = New-Object System.Drawing.Bitmap($dst)
    Write-Output ("{0} <- {1} ({2}x{3})" -f (Split-Path $k -Leaf), (Split-Path $src -Leaf), $b.Width, $b.Height)
    $b.Dispose()
}

# --- правка wp:extent / a:ext: пропорции нового изображения, вписанные в прежний блок ---
$docPath = Join-Path $tmp 'word\document.xml'
$doc = New-Object System.Xml.XmlDocument
$doc.PreserveWhitespace = $true
$doc.Load($docPath)
$ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
$ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
$ns.AddNamespace('wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing')
$ns.AddNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
$ns.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
$ns.AddNamespace('pic', 'http://schemas.openxmlformats.org/drawingml/2006/picture')

$relsText = [System.IO.File]::ReadAllText((Join-Path $tmp 'word\_rels\document.xml.rels'), [System.Text.Encoding]::UTF8)
$relMap = @{}
foreach ($m in [regex]::Matches($relsText, '<Relationship\s[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"')) {
    $relMap[$m.Groups[1].Value] = $m.Groups[2].Value
}

$newRatio = @{}
foreach ($k in $map.Keys) {
    $name = Split-Path $k -Leaf
    $bmp = New-Object System.Drawing.Bitmap((Join-Path $tmp $k))
    $newRatio[$name] = $bmp.Width / $bmp.Height
    $bmp.Dispose()
}

foreach ($d in @($doc.SelectNodes('//w:drawing', $ns))) {
    $blip = $d.SelectSingleNode('.//a:blip', $ns)
    if ($null -eq $blip) { continue }
    $rid = $blip.GetAttribute('embed', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    if (-not $relMap.ContainsKey($rid)) { continue }
    $name = Split-Path ($relMap[$rid] -replace '/', '\') -Leaf
    if (-not $newRatio.ContainsKey($name)) { continue }
    $ratio = $newRatio[$name]
    $ext = $d.SelectSingleNode('.//wp:extent', $ns)
    if ($null -eq $ext) { continue }
    $cx = [int64]$ext.GetAttribute('cx')
    $cy = [int64]$ext.GetAttribute('cy')
    # вписываем с сохранением пропорций в прежний блок (блок не увеличивается)
    if (($cx / $cy) -gt $ratio) { $newCx = [int64][Math]::Round($cy * $ratio); $newCy = $cy }
    else { $newCx = $cx; $newCy = [int64][Math]::Round($cx / $ratio) }
    $ext.SetAttribute('cx', $newCx.ToString())
    $ext.SetAttribute('cy', $newCy.ToString())
    foreach ($ae in @($d.SelectNodes('.//pic:spPr/a:xfrm/a:ext', $ns))) {
        $ae.SetAttribute('cx', $newCx.ToString())
        $ae.SetAttribute('cy', $newCy.ToString())
    }
    Write-Output ("{0} ({1}): extent {2}x{3} -> {4}x{5}" -f $name, $rid, $cx, $cy, $newCx, $newCy)
}

$settings = New-Object System.Xml.XmlWriterSettings
$settings.Encoding = New-Object System.Text.UTF8Encoding($false)
$settings.Indent = $false
$settings.OmitXmlDeclaration = $false
$wtr = [System.Xml.XmlWriter]::Create($docPath, $settings)
$doc.Save($wtr)
$wtr.Close()

# --- упаковка в каноническом порядке ---
$outTmp = Join-Path $env:TEMP ("docx_colorfix_out_$stamp.docx")
$all = [System.IO.Directory]::GetFiles($tmp, '*', [System.IO.SearchOption]::AllDirectories)
$rel = @()
foreach ($f in $all) { $rel += ($f.Substring($tmp.Length + 1) -replace '\\', '/') }
$ordered = @()
$ordered += ($rel | Where-Object { $_ -eq '[Content_Types].xml' })
$ordered += ($rel | Where-Object { $_ -eq '_rels/.rels' })
$ordered += ($rel | Where-Object { $_ -ne '[Content_Types].xml' -and $_ -ne '_rels/.rels' -and $_ -notlike 'word/media/*' } | Sort-Object)
$ordered += ($rel | Where-Object { $_ -like 'word/media/*' } | Sort-Object)
$fs = [System.IO.File]::Open($outTmp, [System.IO.FileMode]::CreateNew)
$zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
foreach ($name in $ordered) {
    $src = Join-Path $tmp ($name -replace '/', '\')
    $entry = $zip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
    $es = $entry.Open()
    $bytes = [System.IO.File]::ReadAllBytes($src)
    $es.Write($bytes, 0, $bytes.Length)
    $es.Close()
}
$zip.Dispose()
$fs.Close()

# проверка пакета до подмены
$z = [System.IO.Compression.ZipFile]::OpenRead($outTmp)
$names = @($z.Entries | ForEach-Object { $_.FullName })
$media = @($names | Where-Object { $_ -like 'word/media/*' })
$ink = @($names | Where-Object { $_ -like 'word/ink/*' })
$z.Dispose()
Write-Output ("пакет: записей {0}, media {1}, ink {2}" -f $names.Count, $media.Count, $ink.Count)
if ($media.Count -ne 25) { throw "неверное число media: $($media.Count)" }

Copy-Item -LiteralPath $outTmp -Destination $work -Force
Write-Output ("рабочий файл обновлён: {0} байт" -f (Get-Item -LiteralPath $work).Length)

try { Remove-Item -LiteralPath $tmp -Recurse -Force } catch { Write-Output "временная папка не удалена: $($_.Exception.Message)" }
try { Remove-Item -LiteralPath $outTmp -Force } catch { }
