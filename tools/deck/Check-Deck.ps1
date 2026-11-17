param(
    [string]$ProjectRoot = '.',
    [string]$File = 'Презентация_курсовой_проект_Копейкин.pptx'
)
$ErrorActionPreference = 'Stop'
$path = Join-Path $ProjectRoot $File
$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = -1
$pres = $ppt.Presentations.Open($path, $true, $false, $false)
"Слайдов: " + $pres.Slides.Count
"Размер: " + $pres.PageSetup.SlideWidth + " x " + $pres.PageSetup.SlideHeight + " pt"
foreach ($s in $pres.Slides) {
    $title = ''
    $pics = 0
    $tables = 0
    $shapes = 0
    $note = ''
    foreach ($sh in $s.Shapes) {
        $shapes++
        if ($sh.Type -eq 13) { $pics++ }
        if ($sh.HasTable -eq -1) { $tables++ }
    }
    # первый текстовый блок с наибольшим размером шрифта — заголовок
    $best = 0; $bestTxt = ''
    foreach ($sh in $s.Shapes) {
        if ($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1) {
            $t = $sh.TextFrame.TextRange.Text
            if ([string]::IsNullOrWhiteSpace($t)) { continue }
            $sz = $sh.TextFrame.TextRange.Font.Size
            if ($sz -gt $best) { $best = $sz; $bestTxt = ($t -replace "[\r\n\x07]", ' ') }
        }
    }
    if ($bestTxt.Length -gt 70) { $bestTxt = $bestTxt.Substring(0, 70) + '…' }
    $notesTxt = ''
    if ($s.NotesPage.Shapes.Count -ge 2) {
        $notesTxt = ($s.NotesPage.Shapes.Item(2).TextFrame.TextRange.Text -replace "[\r\n]", ' ')
    }
    if ($notesTxt.Length -gt 50) { $notesTxt = $notesTxt.Substring(0, 50) + '…' }
    "[{0}] shapes={1} pics={2} tables={3} | {4}" -f $s.SlideIndex, $shapes, $pics, $tables, $bestTxt
    if ($notesTxt -ne '') { "      notes: " + $notesTxt }
}
$pres.Close()
$ppt.Quit()
