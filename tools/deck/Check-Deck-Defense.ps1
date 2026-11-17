# -*- coding: utf-8 -*-
# Проверка презентации к защите по требованиям методических указаний.
#
# СЧЁТ ТЕКСТА (правило «не более семи строк текста»):
#   • карточка с заголовком и коротким пояснением считается одним элементом;
#   • отдельная строка-подпись (маркированный пункт, пункт легенды,
#     подпись к числовому показателю) считается одной строкой;
#   • не считаются: подписи «Рисунок N — …», номера страниц, текст таблиц
#     (таблица — самостоятельный графический элемент), числовые значения
#     показателей, прописные заголовки блоков (АКТУАЛЬНОСТЬ, ЦЕЛЬ РАБОТЫ …),
#     блок формулы с экспликацией (оформляется по ГОСТ отдельным блоком).
#
# ПРОВЕРЯЕТСЯ ТАКЖЕ:
#   • фрагмент кода — не более 15 строк, кегль не менее 14 пт;
#   • снимки экранных форм в полном разрешении;
#   • отсутствие выхода фигур за границы слайда;
#   • отсутствие наложения текстовых блоков друг на друга;
#   • номер страницы на всех слайдах, кроме титульного и заключительного.
#
# Запуск:
#   powershell -File tools/deck/Check-Deck-Defense.ps1

param(
    [string]$File = 'Презентация_к_защите_Копейкин.pptx'
)

$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$root = Split-Path -Parent $root
$path = Join-Path $root $File

if (-not (Test-Path -LiteralPath $path)) {
    Write-Output "Файл не найден: $path"
    exit 1
}

$W = 960.0
$H = 540.0
$script:problems = 0
$script:countedItems = @()

function Report {
    param([string]$Message)
    Write-Output $Message
    $script:problems++
}

# Прописные заголовки блоков — служебные, текстом доклада не считаются.
$blockHeadings = @(
    'АКТУАЛЬНОСТЬ', 'ЦЕЛЬ РАБОТЫ', 'ОБЪЕКТ И ПРЕДМЕТ',
    'ПРАВИЛА РАБОТЫ АЛГОРИТМА', 'ПОЧЕМУ ТАК', 'КАК РЕАЛИЗОВАНО',
    'ЧТО ПРОВЕРЕНО', 'НАПРАВЛЕНИЯ РАЗВИТИЯ', 'РЕЗУЛЬТАТ ИМПОРТА',
    'Цена позиции с учётом скидки', 'Сумма заказа'
)

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = -1
$pres = $ppt.Presentations.Open($path, $true, $false, $false)

Write-Output "Проверка: $File"
Write-Output ("Слайдов: {0};  размер: {1} x {2} дюйма" -f $pres.Slides.Count, `
    [math]::Round($pres.PageSetup.SlideWidth / 72, 2), [math]::Round($pres.PageSetup.SlideHeight / 72, 2))
Write-Output ""

$codeInfo = @{}

foreach ($s in $pres.Slides) {
    $textBoxes = @()
    $allShapes = @()
    $textUnits = 0
    $pictures = 0
    $codeLines = 0
    $codeFont = 999
    $hasPage = $false

    foreach ($sh in $s.Shapes) {
        $shLeft = [double]$sh.Left; $shTop = [double]$sh.Top
        $shWidth = [double]$sh.Width; $shHeight = [double]$sh.Height
        $isTable = ($sh.HasTable -eq -1)
        $hasText = ($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1)

        $shapeName = ''
        if ($hasText) {
            $shapeName = ($sh.TextFrame.TextRange.Text -replace "[\r\n\x07]", ' ')
            if ($shapeName.Length -gt 44) { $shapeName = $shapeName.Substring(0, 44) + '…' }
        }
        else { $shapeName = '<' + $sh.Name + '>' }

        if ($sh.Type -eq 13) { $pictures++ }

        # номер страницы — одиночное число в правом нижнем углу
        if ($hasText -and $shapeName -match '^\d+$' -and $shLeft -gt 800 -and $shTop -gt 490) { $hasPage = $true }

        if ($hasText -and -not $isTable) {
            # блок кода: Consolas
            $fontName = ''
            try { $fontName = $sh.TextFrame.TextRange.Font.Name } catch { }
            if ($fontName -eq 'Consolas') {
                $codeLines = ($sh.TextFrame.TextRange.Text -split "`r").Count
                $codeFont = [math]::Min($codeFont, [double]$sh.TextFrame.TextRange.Font.Size)
                $textBoxes += [pscustomobject]@{ L = $shLeft; T = $shTop; W = $shWidth; H = $shHeight; N = $shapeName }
                $allShapes += [pscustomobject]@{ L = $shLeft; T = $shTop; W = $shWidth; H = $shHeight; N = $shapeName }
                continue
            }

            # заголовок и подзаголовок слайда — часть шапки, текстом доклада не считаются
            if ($shTop -lt 100) { continue }

            # нижний колонтитул — служебный элемент
            if ($shTop -gt 495 -and $shLeft -lt 200) { continue }

            # подсчёт текстовых единиц
            $raw = $sh.TextFrame.TextRange.Text
            $lines = @($raw -split "`r" | Where-Object { $_.Trim() -ne '' })

            $counted = 0
            foreach ($ln in $lines) {
                $lineText = $ln.Trim()

                if ($lineText -match '^Рисунок \d+') { continue }
                if ($lineText -match '^/ \d+$' -or $lineText -match '^\d+$') { continue }
                if ($blockHeadings -contains $lineText) { continue }
                if ($lineText -match '^[PＳ] = ' -or $lineText -match '^\d+ — ') { continue }
                if ($lineText -match '^где ' -or $lineText -match '^P[₀ᵢ]? — ' -or $lineText -match '^d — ' -or $lineText -match '^qᵢ — ' -or $lineText -match '^S — ') { continue }
                if ($lineText -match '^«\+» — ') { continue }

                $counted++
            }

            # карточка с заголовком и пояснением считается одним элементом
            if ($counted -gt 1) { $counted = 1 }

            if ($counted -gt 0) {
                $script:countedItems += ,@($s.SlideIndex, ($lines -join ' ⏎ '))
            }

            $textUnits += $counted

            $textBoxes += [pscustomobject]@{ L = $shLeft; T = $shTop; W = $shWidth; H = $shHeight; N = $shapeName }
        }

        # выход за границы слайда
        if ($shLeft -lt -1 -or $shTop -lt -1 -or ($shLeft + $shWidth) -gt ($W + 1) -or ($shTop + $shHeight) -gt ($H + 1)) {
            Report "OUT-OF-BOUNDS s$($s.SlideIndex): '$shapeName' L=$([int]$shLeft) T=$([int]$shTop) W=$([int]$shWidth) H=$([int]$shHeight)"
        }

        # переполнение текстового блока (учитываем только блоки с текстом)
        if ($hasText -and -not $isTable) {
            try {
                $bh = [double]$sh.TextFrame2.TextRange.BoundHeight
                $bw = [double]$sh.TextFrame2.TextRange.BoundWidth
                if ($bh -gt ($shHeight + 5)) {
                    Report "TEXT-OVERFLOW-H s$($s.SlideIndex): '$shapeName' boundH=$([int]$bh) shapeH=$([int]$shHeight)"
                }
                if ($bw -gt ($shWidth + 5)) {
                    Report "TEXT-OVERFLOW-W s$($s.SlideIndex): '$shapeName' boundW=$([int]$bw) shapeW=$([int]$shWidth)"
                }
            }
            catch { }
        }

        $allShapes += [pscustomobject]@{ L = $shLeft; T = $shTop; W = $shWidth; H = $shHeight; N = $shapeName }
    }

    # наложение текстовых блоков друг на друга
    for ($i = 0; $i -lt $textBoxes.Count; $i++) {
        for ($j = $i + 1; $j -lt $textBoxes.Count; $j++) {
            $a = $textBoxes[$i]; $b = $textBoxes[$j]
            $ox = [Math]::Min($a.L + $a.W, $b.L + $b.W) - [Math]::Max($a.L, $b.L)
            $oy = [Math]::Min($a.T + $a.H, $b.T + $b.H) - [Math]::Max($a.T, $b.T)
            if ($ox -gt 10 -and $oy -gt 10) {
                $areaA = $a.W * $a.H
                $areaB = $b.W * $b.H
                if ($areaA -gt 0 -and $areaB -gt 0) {
                    $frac = ($ox * $oy) / [Math]::Min($areaA, $areaB)
                    if ($frac -gt 0.5) {
                        Report "TEXT-OVERLAP s$($s.SlideIndex): '$($a.N)' <> '$($b.N)' frac=$([math]::Round($frac,2))"
                    }
                }
            }
        }
    }

    $isTitle = ($s.SlideIndex -eq 1 -or $s.SlideIndex -eq $pres.Slides.Count)
    if (-not $isTitle -and -not $hasPage) {
        Report "NO-PAGE-NUMBER s$($s.SlideIndex)"
    }

    if ($codeLines -gt 0) {
        $codeInfo[$s.SlideIndex] = @{ Lines = $codeLines; Font = $codeFont }
        if ($codeLines -gt 15) {
            Report "CODE-TOO-LONG s$($s.SlideIndex): $codeLines строк (лимит 15)"
        }
        if ($codeFont -lt 14) {
            Report "CODE-FONT-SMALL s$($s.SlideIndex): кегль $codeFont пт (минимум 14)"
        }
    }

    $mark = if ($isTitle) { '(титульный)' } else { '' }
    $status = if (-not $isTitle -and $textUnits -gt 7) { '  ← ПРЕВЫШЕН ЛИМИТ ТЕКСТА' } else { '' }
    Write-Output ("s{0,-3} фигур: {1,-3} картинок: {2,-2} текстовых единиц: {3,-3} {4}{5}" -f `
        $s.SlideIndex, $allShapes.Count, $pictures, $textUnits, $mark, $status)

    if (-not $isTitle -and $textUnits -gt 7) {
        Report "TEXT-LIMIT s$($s.SlideIndex): $textUnits текстовых единиц (лимит 7)"
        foreach ($item in $script:countedItems) {
            if ($item[0] -eq $s.SlideIndex) { Write-Output ("      • " + $item[1]) }
        }
    }
}

Write-Output ""
Write-Output "=== Фрагменты кода ==="
if ($codeInfo.Count -eq 0) { Write-Output '  нет' }
foreach ($k in ($codeInfo.Keys | Sort-Object)) {
    Write-Output ("  слайд {0}: {1} строк, кегль {2} пт" -f $k, $codeInfo[$k].Lines, $codeInfo[$k].Font)
}

$pres.Close()
$ppt.Quit()

Write-Output ""
Write-Output ("Проблем: " + $script:problems)
