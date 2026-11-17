<#
.SYNOPSIS
    Заполняет текстовый плейсхолдер рядом со схемой архитектуры
    и уточняет подпись рисунка в пояснительной записке.

.DESCRIPTION
    1. Делает резервную копию .docx.
    2. Заменяет «ВАШ ТЕКСТ…» (тот, что стоит перед фразой «Схема
       взаимодействия слоев приложения приведена на рисунке 2.»)
       на описание архитектуры приложения.
    3. Заменяет подпись «Рисунок 2 – Пример схемы взаимодействия слоев»
       на «Рисунок 2 – Архитектура ИС «Магазин автозапчастей»».
#>
param(
    [string]$Docx = 'Копейкин _пояснительная_записка_курсового_проекта.docx',
    [string]$ProjectRoot = '.'
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

$docxPath = Join-Path $ProjectRoot $Docx
$backup = Join-Path (Join-Path $ProjectRoot 'docs') ('backup_zapiska_' + (Get-Date -Format 'yyyy-MM-dd-HHmmss') + '.docx')

Copy-Item -LiteralPath $docxPath -Destination $backup -Force
Write-Output "Резервная копия: $backup"

$zip = [System.IO.Compression.ZipFile]::Open($docxPath, 'Update')
try {
    $entry = $zip.GetEntry('word/document.xml')
    $sr = New-Object System.IO.StreamReader($entry.Open(), [Text.Encoding]::UTF8)
    $xml = $sr.ReadToEnd()
    $sr.Close()

    # 1. Плейсхолдер перед фразой про рисунок 2
    $anchor = 'Схема взаимодействия слоев приложения приведена на рисунке 2.'
    $anchorIdx = $xml.IndexOf($anchor)
    if ($anchorIdx -lt 0) { throw 'Не найдена фраза «Схема взаимодействия слоев…»' }

    $ph = 'ВАШ ТЕКСТ…'
    $phIdx = $xml.LastIndexOf($ph, $anchorIdx)
    if ($phIdx -lt 0) { throw 'Не найден плейсхолдер «ВАШ ТЕКСТ…» перед схемой' }

    $newText = 'Архитектура приложения представлена на рисунке 2. Программа разделена на четыре уровня: интерфейс пользователя (Avalonia), логика окон (ViewModel), бизнес-правила (сервисы) и доступ к данным (контекст Entity Framework Core). Каждый уровень обращается только к следующему, поэтому замена хранилища или правил не затрагивает интерфейс.'

    $xml = $xml.Remove($phIdx, $ph.Length).Insert($phIdx, $newText)
    Write-Output 'OK плейсхолдер «ВАШ ТЕКСТ…» заменён'

    # 2. Подпись рисунка 2
    $oldCap = 'Рисунок 2 – Пример схемы взаимодействия слоев'
    $newCap = 'Рисунок 2 – Архитектура ИС «Магазин автозапчастей»'
    if ($xml.Contains($oldCap)) {
        $xml = $xml.Replace($oldCap, $newCap)
        Write-Output 'OK подпись рисунка 2 уточнена'
    }
    else {
        Write-Output 'ВНИМАНИЕ: подпись «Рисунок 2 – Пример схемы взаимодействия слоев» не найдена'
    }

    $entry.Delete()
    $newEntry = $zip.CreateEntry('word/document.xml', [System.IO.Compression.CompressionLevel]::Optimal)
    $sw = New-Object System.IO.StreamWriter($newEntry.Open(), (New-Object Text.UTF8Encoding($false)))
    $sw.Write($xml)
    $sw.Close()
}
finally {
    $zip.Dispose()
}

Write-Output 'Готово.'
