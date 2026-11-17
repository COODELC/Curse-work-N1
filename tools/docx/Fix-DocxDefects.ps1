$ErrorActionPreference = 'Stop'
$root = 'F:\курсовая Копейкин Андрей\Копейкин Андрей\Копейкин Андрей Последняя версия — копия'
$doc = Join-Path $root 'Копейкин _пояснительная_записка_курсового_проекта.docx'
$stamp = Get-Date -Format 'yyyy-MM-dd-HHmmss'
$bak = Join-Path $root ('docs\backup_zapiska_' + $stamp + '.docx')
Copy-Item -LiteralPath $doc -Destination $bak -Force
Write-Output ('Бэкап: ' + $bak)

$w = New-Object -ComObject Word.Application
$w.Visible = $false; $w.DisplayAlerts = 0
$d = $w.Documents.Open($doc)

function Get-PlainText($p) {
    return (($p.Range.Text -replace "[`r`n`v`a`t]", ' ')).Trim()
}

# ---------- сборка списка абзацев ----------
$lines = New-Object System.Collections.ArrayList
$i = 0
foreach ($p in $d.Paragraphs) {
    $i++
    [void]$lines.Add([pscustomobject]@{ N = $i; T = (Get-PlainText $p) })
}
$arr = $lines.ToArray()
$count = $arr.Count

function Get-NextNonEmpty([int]$from) {
    for ($j = $from; $j -le $count; $j++) {
        if ($arr[$j - 1].T -ne '') { return $arr[$j - 1].T }
    }
    return ''
}
function Get-PrevNonEmpty([int]$from) {
    for ($j = $from; $j -ge 1; $j--) {
        if ($arr[$j - 1].T -ne '') { return $arr[$j - 1].T }
    }
    return ''
}

# ---------- 1. Удаление дублирующихся подписей «Таблица N» ----------
# Удаляем «Таблица N», только если рядом (до или после) стоит «Таблица N — Название»
$targets = New-Object System.Collections.ArrayList
foreach ($l in $arr) {
    if ($l.T -match '^Таблица\s+(\d+)$') {
        $num = $matches[1]
        $nx = Get-NextNonEmpty ($l.N + 1)
        $pv = Get-PrevNonEmpty ($l.N - 1)
        if (($nx -match ('^Таблица\s+' + $num + '\s+[—–-]')) -or ($pv -match ('^Таблица\s+' + $num + '\s+[—–-]'))) {
            [void]$targets.Add($l.N)
            Write-Output ('  дубль: абз ' + $l.N + ' = [' + $l.T + '] (сосед: [' + $(if ($nx -match ('^Таблица\s+' + $num)) { $nx } else { $pv }) + '])')
        }
    }
}
Write-Output ('К удалению: ' + $targets.Count)

$removed = 0
foreach ($n in ($targets | Sort-Object -Descending)) {
    $p = $d.Paragraphs.Item($n)
    if ((Get-PlainText $p) -match '^Таблица\s+\d+$') {
        [void]$p.Range.Delete()
        $removed++
    }
}
Write-Output ('Удалено: ' + $removed)

# ---------- 2. Латинская «A» → кириллическая «А» ----------
$cyr = [string][char]1040
$fixed = 0
$i = 0
foreach ($p in $d.Paragraphs) {
    $i++
    $t = Get-PlainText $p
    if ($t -match 'ПРИЛОЖЕНИЕ\s+A' -or $t -match '^Рисунок\s+A\.1') {
        $r = $p.Range
        $f = $r.Find
        $f.ClearFormatting()
        $f.Text = 'A'
        $f.Replacement.ClearFormatting()
        $f.Replacement.Text = $cyr
        $f.Forward = $true
        $f.Wrap = 0
        $f.MatchCase = $true
        $f.MatchWholeWord = $false
        $f.MatchWildcards = $false
        $f.MatchSoundsLike = $false
        $f.MatchAllWordForms = $false
        $f.Format = $false
        $null = $f.Execute([ref]'A', [ref]$true, [ref]$false, [ref]$false, [ref]$false, [ref]$false,
                            [ref]$true, [ref]0, [ref]$false, [ref]$cyr, [ref]2)
        $fixed++
        Write-Output ('  абз ' + $i + ' -> ' + (Get-PlainText $p))
    }
}
Write-Output ('Исправлено абзацев с латинской A: ' + $fixed)

# ---------- 3. Подпись рис. 3 — в отдельный абзац ----------
$done = $false
$i = 0
foreach ($p in $d.Paragraphs) {
    $i++
    $t = Get-PlainText $p
    if ($t -match '^Рисунок\s*3\s*—' -and $p.Range.InlineShapes.Count -gt 0) {
        $r = $p.Range.Duplicate
        $f = $r.Find
        $f.ClearFormatting()
        $f.Text = 'Рисунок 3 —'
        $f.Forward = $true
        $f.Wrap = 0
        $f.MatchCase = $true
        $f.MatchWholeWord = $false
        $f.MatchWildcards = $false
        if ($f.Execute()) {
            $r.InsertParagraphBefore()
            $done = $true
            Write-Output ('  подпись рис. 3 вынесена из абзаца ' + $i)
        }
        break
    }
}
Write-Output ('Рис.3 вынесен: ' + $done)

# ---------- 4. Обновление оглавления и полей ----------
foreach ($toc in $d.TablesOfContents) { $toc.Update() }
$d.Repaginate()
$d.Save()
$pages = $d.ComputeStatistics(2)
$words = $d.ComputeStatistics(0)
$d.Close($false)
$w.Quit()
Write-Output ("OK сохранено. Страниц: $pages, слов: $words")
