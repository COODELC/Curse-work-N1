$ErrorActionPreference = 'Stop'
$root = 'F:\курсовая Копейкин Андрей\Копейкин Андрей\Копейкин Андрей Последняя версия — копия'
$docPath = Join-Path $root 'Копейкин _пояснительная_записка_курсового_проекта.docx'

$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
$w.Options.UpdateFieldsAtPrint = $false
$w.Options.Pagination = $true
Write-Output 'Word запущен'
$d = $w.Documents.Open($docPath, $false, $false)
Write-Output 'Документ открыт'

function Get-PText($p) {
    return (($p.Range.Text -replace "[`r`n`v`a`t]", ' ')).Trim()
}

$caps = @()
foreach ($p in $d.Paragraphs) {
    $t = Get-PText $p
    if ($t -match '^(Рисунок|Таблица)\s') { $caps += $t }
}
Write-Output ('Подписей всего: ' + $caps.Count)
$dup = $caps | Group-Object | Where-Object { $_.Count -gt 1 }
Write-Output ('Дублей подписей: ' + @($dup).Count)

$latinA = 0
$i = 0
foreach ($p in $d.Paragraphs) {
    $i++
    $t = Get-PText $p
    if ($t -match 'ПРИЛОЖЕНИЕ\s+A' -or $t -match 'Рисунок\s+A\.1') {
        $latinA++
        Write-Output ('  ЛАТИНИЦА: ' + $t)
    }
}
Write-Output ('Осталось латинских A: ' + $latinA)

Write-Output ('Страниц до обновления: ' + $d.ComputeStatistics(2))
foreach ($toc in $d.TablesOfContents) { $toc.Update() }
Write-Output 'Оглавление обновлено'
$d.Save()
Write-Output ('Страниц после: ' + $d.ComputeStatistics(2) + '; слов: ' + $d.ComputeStatistics(0) + '; таблиц: ' + $d.Tables.Count + '; рисунков: ' + $d.InlineShapes.Count)
$d.Close($false)
$w.Quit()
Write-Output 'OK сохранено'
