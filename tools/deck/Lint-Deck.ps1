param(
    [string]$ProjectRoot = '.',
    [string]$File = 'Презентация_курсовой_проект_Копейкин.pptx'
)
$ErrorActionPreference = 'Stop'
$path = Join-Path $ProjectRoot $File
$W = 960.0; $H = 540.0
$problems = 0

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = -1
$pres = $ppt.Presentations.Open($path, $true, $false, $false)

foreach ($s in $pres.Slides) {
    $boxes = @()
    foreach ($sh in $s.Shapes) {
        $L = [double]$sh.Left; $T = [double]$sh.Top
        $Wsh = [double]$sh.Width; $Hsh = [double]$sh.Height
        $name = ''
        if ($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1) {
            $name = ($sh.TextFrame.TextRange.Text -replace "[\r\n\x07]", ' ')
            if ($name.Length -gt 40) { $name = $name.Substring(0, 40) + '…' }
        }
        else { $name = '<' + $sh.Name + '>' }

        # выход за пределы слайда
        if ($L -lt -1 -or $T -lt -1 -or ($L + $Wsh) -gt ($W + 1) -or ($T + $Hsh) -gt ($H + 1)) {
            "OUT-OF-BOUNDS s$($s.SlideIndex): '$name' L=$([int]$L) T=$([int]$T) W=$([int]$Wsh) H=$([int]$Hsh)"
            $problems++
        }
        # переполнение текстового блока
        if ($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1) {
            try {
                $bh = [double]$sh.TextFrame2.TextRange.BoundHeight
                $bw = [double]$sh.TextFrame2.TextRange.BoundWidth
                if ($bh -gt ($Hsh + 3)) {
                    "TEXT-OVERFLOW-H s$($s.SlideIndex): '$name' boundH=$([int]$bh) shapeH=$([int]$Hsh)"
                    $problems++
                }
                if ($bw -gt ($Wsh + 3)) {
                    "TEXT-OVERFLOW-W s$($s.SlideIndex): '$name' boundW=$([int]$bw) shapeW=$([int]$Wsh)"
                    $problems++
                }
            }
            catch { }
        }
        $boxes += [pscustomobject]@{ L = $L; T = $T; W = $Wsh; H = $Hsh; N = $name }
    }
    # грубая проверка наложения текстовых блоков (только непрозрачные текстовые фигуры)
    for ($i = 0; $i -lt $boxes.Count; $i++) {
        for ($j = $i + 1; $j -lt $boxes.Count; $j++) {
            $a = $boxes[$i]; $b = $boxes[$j]
            $ox = [Math]::Min($a.L + $a.W, $b.L + $b.W) - [Math]::Max($a.L, $b.L)
            $oy = [Math]::Min($a.T + $a.H, $b.T + $b.H) - [Math]::Max($a.T, $b.T)
            if ($ox -gt 8 -and $oy -gt 8) {
                $areaA = $a.W * $a.H
                $areaB = $b.W * $b.H
                $ov = $ox * $oy
                if ($areaA -gt 0 -and $areaB -gt 0) {
                    $frac = $ov / [Math]::Min($areaA, $areaB)
                    if ($frac -gt 0.35) {
                        "OVERLAP s$($s.SlideIndex): '$($a.N)' <> '$($b.N)' frac=$([math]::Round($frac,2))"
                        $problems++
                    }
                }
            }
        }
    }
}
$pres.Close()
$ppt.Quit()
"Проблем: " + $problems
