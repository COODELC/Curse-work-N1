param(
    [string]$ProjectRoot = '.',
    [string]$File = 'Презентация_курсовой_проект_Копейкин.pptx',
    [string]$OutDir = 'docs\deck_png'
)
$ErrorActionPreference = 'Stop'
$path = Join-Path $ProjectRoot $File
$outDir = Join-Path $ProjectRoot $OutDir
if (-not (Test-Path -LiteralPath $outDir)) {
    New-Item -ItemType Directory -Path $outDir | Out-Null
}
$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = -1
$pres = $ppt.Presentations.Open($path, $true, $false, $false)
foreach ($s in $pres.Slides) {
    $png = Join-Path $outDir ("slide-" + $s.SlideIndex.ToString('00') + ".png")
    $s.Export($png, 'PNG', 1280, 720)
}
$pres.Close()
$ppt.Quit()
Get-ChildItem -LiteralPath $outDir -Filter '*.png' | Measure-Object | Select-Object -ExpandProperty Count
