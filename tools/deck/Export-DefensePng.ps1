# -*- coding: utf-8 -*-
# Экспорт слайдов презентации в PNG для визуальной проверки.
#
# Запуск:
#   powershell -File tools/deck/Export-DefensePng.ps1

param(
    [string]$File = 'Презентация_к_защите_Копейкин.pptx',
    [string]$OutDir = 'docs\deck_defense_png',
    [int]$Width = 1600,
    [int]$Height = 900
)

$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$root = Split-Path -Parent $root

$path = Join-Path $root $File
$outDir = Join-Path $root $OutDir

if (-not (Test-Path -LiteralPath $outDir)) {
    New-Item -ItemType Directory -Force -Path $outDir | Out-Null
}

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = -1
$pres = $ppt.Presentations.Open($path, $true, $false, $false)

foreach ($s in $pres.Slides) {
    $png = Join-Path $outDir ("slide-" + $s.SlideIndex.ToString('00') + ".png")
    $s.Export($png, 'PNG', $Width, $Height)
    Write-Output ("экспортирован: " + $png)
}

$pres.Close()
$ppt.Quit()

Write-Output ("Всего PNG: " + (Get-ChildItem -LiteralPath $outDir -Filter '*.png').Count)
