$ErrorActionPreference = 'Stop'
$root = 'F:\курсовая Копейкин Андрей\Копейкин Андрей\Копейкин Андрей Последняя версия — копия'
$docPath = Join-Path $root 'Копейкин _пояснительная_записка_курсового_проекта.docx'

$w = New-Object -ComObject Word.Application
$w.Visible = $false; $w.DisplayAlerts = 0
$d = $w.Documents.Open($docPath, $false, $true)
Write-Output ('Открыт без восстановления. Страниц: ' + $d.ComputeStatistics(2) + '; слов: ' + $d.ComputeStatistics(0) + '; таблиц: ' + $d.Tables.Count + '; рисунков: ' + $d.InlineShapes.Count + '; абзацев: ' + $d.Paragraphs.Count + '; оглавлений: ' + $d.TablesOfContents.Count)

function Get-PText($p) { return (($p.Range.Text -replace "[`r`n`v`a`t]", ' ')).Trim() }

$figs = @(); $tabs = @(); $dup = 0
foreach ($p in $d.Paragraphs) {
    $t = Get-PText $p
    if ($t -match '^Рисунок\s') { $figs += $t }
    if ($t -match '^Таблица\s') { $tabs += $t }
}
$all = $figs + $tabs
$dup = @($all | Group-Object | Where-Object { $_.Count -gt 1 }).Count
Write-Output ('Подписей рисунков: ' + $figs.Count + '; подписей таблиц: ' + $tabs.Count + '; дубликатов: ' + $dup)
Write-Output '--- подписи рисунков ---'
$figs | ForEach-Object { '  ' + $_ }
Write-Output '--- подписи таблиц ---'
$tabs | ForEach-Object { '  ' + $_ }

Write-Output '--- абзац с подписью рис. 3 ---'
$i = 0
foreach ($p in $d.Paragraphs) {
    $i++
    $t = Get-PText $p
    if ($t -match '^Рисунок\s*3\s*—') {
        $prev = Get-PText $d.Paragraphs.Item($i - 1)
        "  абз $i : [$t]"
        "  предыдущий абз: [$prev]  (изображений в нём: $($d.Paragraphs.Item($i-1).Range.InlineShapes.Count))"
        "  изображений в абзаце подписи: $($p.Range.InlineShapes.Count)"
    }
}
Write-Output '--- заголовки приложений ---'
foreach ($p in $d.Paragraphs) {
    if ($p.Style.NameLocal -eq 'Заголовок 1') {
        $t = Get-PText $p
        if ($t -like 'ПРИЛОЖЕНИЕ*') {
            $ch = $t.Substring(10, 1)
            "  '$t' | буква код=$([int][char]$ch) $(if([int][char]$ch -ge 1040){'(кириллица)'}else{'(ЛАТИНИЦА!)'})"
        }
    }
}
$d.Close($false); $w.Quit()
Write-Output 'OK'
