<#
.SYNOPSIS
    Заменяет рисунки в пояснительной записке: схему слоёв и схему архитектуры.

.DESCRIPTION
    1. Делает резервную копию .docx.
    2. Подменяет word/media/image3.png (схема «формы — сервисы — база»)
       на docs/diagrams/datastreams.png.
    3. Подменяет word/media/image8.png (схема архитектуры) на
       docs/diagrams/architecture.png.
    4. Пропорционально пересчитывает wp:extent и a:ext, чтобы рисунок
       не исказился.
    5. Заменяет текст-плейсхолдер «ВАШ ТЕКСТ…» перед схемой и подпись
       «Рисунок 2 – Пример схемы взаимодействия слоев».
#>
param(
    [string]$Docx = 'Копейкин _пояснительная_записка_курсового_проекта.docx',
    [string]$ProjectRoot = '.'
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

$docxPath = Join-Path $ProjectRoot $Docx
$backupDir = Join-Path $ProjectRoot 'docs'
$stamp = Get-Date -Format 'yyyy-MM-dd-HHmmss'
$backup = Join-Path $backupDir ("backup_zapiska_" + $stamp + ".docx")

if (-not (Test-Path -LiteralPath $docxPath)) { throw "Не найден документ: $docxPath" }

Copy-Item -LiteralPath $docxPath -Destination $backup -Force
Write-Output "Резервная копия: $backup"

# Целевые рисунки: media -> файл, ratio берётся из самого PNG
$targets = @(
    @{ Media = 'word/media/image3.png'; File = 'docs/diagrams/datastreams.png' },
    @{ Media = 'word/media/image8.png'; File = 'docs/diagrams/architecture.png' }
)

$zip = [System.IO.Compression.ZipFile]::Open($docxPath, 'Update')
try {
    foreach ($t in $targets) {
        $src = Join-Path $ProjectRoot ($t.File -replace '\\','/')
        if (-not (Test-Path -LiteralPath $src)) { throw "Нет исходного рисунка: $src" }

        $entry = $zip.GetEntry($t.Media)
        if ($null -eq $entry) { throw "В документе нет записи $($t.Media)" }

        # Текущий размер рисунка в документе (EMU)
        $docEntry = $zip.GetEntry('word/document.xml')
        $sr = New-Object System.IO.StreamReader($docEntry.Open(), [Text.Encoding]::UTF8)
        $xml = $sr.ReadToEnd()
        $sr.Close()

        $mediaName = Split-Path $t.Media -Leaf
        $relId = ([regex]::Match($xml, 'r:embed="(rId\d+)"(?=[^>]*>)')).Groups
        # находим rId картинки по имени файла в document.xml.rels
        $relsEntry = $zip.GetEntry('word/_rels/document.xml.rels')
        $sr2 = New-Object System.IO.StreamReader($relsEntry.Open(), [Text.Encoding]::UTF8)
        $rels = $sr2.ReadToEnd()
        $sr2.Close()

        $relMatch = [regex]::Match($rels, ('Id="(rId\d+)"[^>]*Target="media/' + [regex]::Escape($mediaName) + '"'))
        if (-not $relMatch.Success) {
            $relMatch = [regex]::Match($rels, ('Target="media/' + [regex]::Escape($mediaName) + '"[^>]*Id="(rId\d+)"'))
        }
        if (-not $relMatch.Success) { throw "Не найдена связь для $mediaName" }
        $rid = $relMatch.Groups[1].Value

        # Размер в документе для этой rId
        $extMatch = [regex]::Match($xml, '<wp:extent cx="(\d+)" cy="(\d+)"[^>]*/>(?:(?!</w:drawing>).)*?r:embed="' + $rid + '"', 'Singleline')
        if (-not $extMatch.Success) {
            $extMatch = [regex]::Match($xml, 'r:embed="' + $rid + '"', 'Singleline')
        }
        $docExt = [regex]::Match($xml, '<wp:extent cx="(\d+)" cy="(\d+)"[^>]*/>(?:(?!r:embed=).)*?r:embed="' + $rid + '"', 'Singleline')
        if (-not $docExt.Success) { throw "Не найден wp:extent для $rid" }

        $oldCx = [int64]$docExt.Groups[1].Value
        $oldCy = [int64]$docExt.Groups[2].Value

        # Пропорции нового рисунка
        Add-Type -AssemblyName System.Drawing
        $img = [System.Drawing.Image]::FromFile($src)
        $ratio = $img.Width / $img.Height
        $img.Dispose()

        $newCx = $oldCx
        $newCy = [int64][Math]::Round($oldCx / $ratio)

        # Заменяем размеры во всех местах, где встречается эта пара
        $pair = 'cx="' + $oldCx + '" cy="' + $oldCy + '"'
        $newPair = 'cx="' + $newCx + '" cy="' + $newCy + '"'
        $count = ([regex]::Matches($xml, [regex]::Escape($pair))).Count
        if ($count -eq 0) { throw "Не найдено ни одной пары размеров $pair" }
        $xml = $xml.Replace($pair, $newPair)

        # Подменяем содержимое рисунка
        $entry.Delete()
        $newEntry = $zip.CreateEntry($t.Media, [System.IO.Compression.CompressionLevel]::Optimal)
        $in = [System.IO.File]::OpenRead($src)
        $out = $newEntry.Open()
        $in.CopyTo($out)
        $out.Close()
        $in.Close()

        # Записываем обновлённый document.xml
        $docEntry.Delete()
        $docEntry2 = $zip.CreateEntry('word/document.xml', [System.IO.Compression.CompressionLevel]::Optimal)
        $sw = New-Object System.IO.StreamWriter($docEntry2.Open(), (New-Object Text.UTF8Encoding($false)))
        $sw.Write($xml)
        $sw.Close()

        Write-Output ("OK {0} -> {1}: {2}x{3} EMU (заменено размеров: {4})" -f $mediaName, (Split-Path $t.File -Leaf), $newCx, $newCy, $count)
    }
}
finally {
    $zip.Dispose()
}

Write-Output 'Готово.'
