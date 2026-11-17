$ErrorActionPreference = 'Stop'
$root = 'F:\курсовая Копейкин Андрей\Копейкин Андрей\Копейкин Андрей Последняя версия — копия'
$doc = Join-Path $root 'Копейкин _пояснительная_записка_курсового_проекта.docx'

$w = New-Object -ComObject Word.Application
$w.Visible = $false; $w.DisplayAlerts = 0
$d = $w.Documents.Open($doc, $false, $true)

$paras = @()
$i = 0
foreach ($p in $d.Paragraphs) {
    $i++
    $t = $p.Range.Text
    # убираем служебный маркер встроенного изображения (chr 47) и переводы строк
    $t = ($t -replace "[`r`n`v`a`t]", ' ').Trim()
    $paras += [pscustomobject]@{ N = $i; T = $t; Style = $p.Style.NameLocal; Start = $p.Range.Start }
}

'=== Расположение всех встроенных изображений и подпись под ними ==='
$inline = @($d.InlineShapes)
foreach ($sh in $inline) {
    $start = $sh.Range.Start
    $hostP = $paras | Where-Object { $_.Start -le $start } | Select-Object -Last 1
    $after = $paras | Where-Object { $_.Start -gt $start -and $_.T -ne '' } | Select-Object -First 1
    $cap = ''
    if ($after -and $after.T -match '^(Рисунок|Таблица)\s') { $cap = $after.T }
    $line = '  img@абз {0,5} | подпись: {1}' -f $hostP.N, $(if ($cap) { $cap } else { '<<< НЕТ >>>' })
    $line
}

'=== Список всех абзацев-подписей ==='
$caps = @()
foreach ($p in $paras) {
    if ($p.T -match '^Рисунок\s+\S+\s*[—–-]') { $caps += $p.T }
}
$caps | ForEach-Object { "  $_" }
"Всего подписей 'Рисунок': $($caps.Count); встроенных изображений: $($inline.Count)"

'=== Латиница / кириллица в обозначениях приложений ==='
foreach ($p in $paras) {
    if ($p.T -match 'ПРИЛОЖЕНИЕ\s+(\S)' -or $p.T -match '^Рисунок\s+(\w)\.') {
        $ch = $matches[1]
        $isCyr = ([int][char]$ch -ge 1040 -and [int][char]$ch -le 1103)
        "  [{0}] '{1}' -> '{2}' код={3} {4}" -f $p.N, $p.Style, $ch, [int][char]$ch, $(if ($isCyr) { 'OK' } else { 'ЛАТИНИЦА!' })
    }
}

$d.Close($false); $w.Quit()
