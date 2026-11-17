param(
    [string]$SourceDir = 'tools\docx\diagrams',
    [string]$OutputDir = 'docs\diagrams'
)

$ErrorActionPreference = 'Stop'

$chrome = 'C:\Program Files\Google\Chrome\Application\chrome.exe'

if (-not (Test-Path -LiteralPath $chrome)) {
    $chrome = 'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'
}

if (-not (Test-Path -LiteralPath $chrome)) {
    throw 'Не найден Chrome или Edge для рендера диаграмм.'
}

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

$files = Get-ChildItem -LiteralPath $SourceDir -Filter '*.html' | Sort-Object Name

# Размер макета каждой диаграммы: имя файла -> ширина x высота окна браузера.
$sizes = @{
    'er'        = @(1600, 1120)
    'layers'    = @(1500, 1450)
    'alg_create' = @(1250, 920)
    'alg_cancel' = @(1250, 700)
    'datastreams' = @(1240, 420)
    'architecture' = @(1050, 1010)
    'project_create' = @(1120, 520)
    'solution_structure' = @(1180, 600)
    'build_output' = @(1180, 470)
    'code_dbfactory' = @(1180, 920)
    'package_efcore' = @(980, 420)
    'testrun_output' = @(1080, 700)
}

foreach ($file in $files) {
    $html = $file.FullName -replace '\\', '/'
    $png = Join-Path (Resolve-Path $OutputDir).Path ($file.BaseName + '.png')
    $pngUrl = $png -replace '\\', '/'

    $window = @(1600, 1200)

    if ($sizes.ContainsKey($file.BaseName)) {
        $window = $sizes[$file.BaseName]
    }

    $chromeArgs = @(
        '--headless=new'
        '--disable-gpu'
        '--hide-scrollbars'
        '--force-device-scale-factor=2'
        '--default-background-color=FFFFFFFF'
        "--screenshot=$pngUrl"
        "--window-size=$($window[0]),$($window[1])"
        "file:///$html"
    )

    $output = & $chrome @chromeArgs 2>&1

    Start-Sleep -Milliseconds 400

    if (Test-Path -LiteralPath $png) {
        $size = (Get-Item -LiteralPath $png).Length
        Write-Output ("OK  {0,-22} {1,9:N0} байт" -f $file.BaseName, $size)
    }
    else {
        Write-Output ("ОШИБКА {0}: PNG не создан" -f $file.BaseName)
        $output | Write-Output
    }
}
