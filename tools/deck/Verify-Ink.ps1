param([string]$ProjectRoot = $PSScriptRoot)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

$root = $ProjectRoot
if (-not (Test-Path (Join-Path $root 'Презентация_курсовой_проект_Копейкин.pptx'))) {
    $root = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
}
$pptx = Join-Path $root 'Презентация_курсовой_проект_Копейкин.pptx'
$outDir = Join-Path $root 'docs\deck_png'
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = -1
$pres = $ppt.Presentations.Open($pptx, $true, $false, $false)

$bad = 0
for ($i = 1; $i -le $pres.Slides.Count; $i++) {
    $png = Join-Path $outDir ('slide-{0:D2}.png' -f $i)
    $pres.Slides.Item($i).Export($png, 'PNG', 1280, 720)
    $bmp = [System.Drawing.Bitmap]::FromFile($png)
    try {
        $minX = 99999; $maxX = -1; $minY = 99999; $maxY = -1
        for ($y = 0; $y -lt $bmp.Height; $y += 1) {
            for ($x = 0; $x -lt $bmp.Width; $x += 2) {
                $c = $bmp.GetPixel($x, $y)
                if ($c.R -lt 240 -or $c.G -lt 240 -or $c.B -lt 240) {
                    if ($x -lt $minX) { $minX = $x }
                    if ($x -gt $maxX) { $maxX = $x }
                    if ($y -lt $minY) { $minY = $y }
                    if ($y -gt $maxY) { $maxY = $y }
                }
            }
        }
        $l = [math]::Round($minX * 960 / 1280)
        $r = [math]::Round($maxX * 960 / 1280)
        $t = [math]::Round($minY * 540 / 720)
        $b = [math]::Round($maxY * 540 / 720)
        $flag = ''
        if ($l -lt 0 -or $t -lt 0 -or $r -gt 961 -or $b -gt 541) { $flag = ' <== ВЫХОД ЗА ХОЛСТ'; $bad++ }
        "s{0:D2}: X {1}..{2}  Y {3}..{4}{5}" -f $i, $l, $r, $t, $b, $flag
    }
    finally { $bmp.Dispose() }
}
"Проблем: $bad"
$pres.Close()
$ppt.Quit()
