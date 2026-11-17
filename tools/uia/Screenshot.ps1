# Снимок экрана окна приложения.
#   powershell -File tools/uia/Screenshot.ps1 -ProcessId 3220 -TitleLike "Каталог" -Output docs\main.png

param(
    [int]$ProcessId = 0,
    [string]$TitleLike = "",
    [string]$Output = "docs\screenshot.png",
    [int]$Width = 0,
    [int]$Height = 0
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

Add-Type @'
using System;
using System.Runtime.InteropServices;

public static class NativeMethods {
    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT r);

    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    public static extern bool MoveWindow(IntPtr hWnd, int x, int y, int w, int h, bool repaint);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left, Top, Right, Bottom; }
}
'@

$root = [System.Windows.Automation.AutomationElement]::RootElement
$walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker

# Ищем окно через UI Automation: имя окна доступно надёжнее, чем заголовок Win32.
# Сначала — среди верхнеуровневых окон, затем — среди вложенных (модальные диалоги).
$handle = [IntPtr]::Zero
$windowName = ""

function Select-Window {
    param([System.Windows.Automation.AutomationElement]$Element)

    try {
        $name = $Element.Current.Name
        $ownerPid = $Element.Current.ProcessId
        $native = $Element.Current.NativeWindowHandle
    } catch {
        return $false
    }

    if ($ProcessId -ne 0 -and $ownerPid -ne $ProcessId) { return $false }
    if ($TitleLike -ne "" -and $name -notlike "*$TitleLike*") { return $false }
    if ($native -eq 0) { return $false }

    $script:handle = [IntPtr]$native
    $script:windowName = $name

    return $true
}

$child = $walker.GetFirstChild($root)

while ($null -ne $child) {
    if (Select-Window -Element $child) { break }

    $child = $walker.GetNextSibling($child)
}

if ($handle -eq [IntPtr]::Zero) {
    # Модальные окна входят в дерево как потомки главного окна приложения.
    $child = $walker.GetFirstChild($root)

    while ($null -ne $child -and $handle -eq [IntPtr]::Zero) {
        try {
            $ownerPid = $child.Current.ProcessId
        } catch {
            $child = $walker.GetNextSibling($child)
            continue
        }

        if ($ProcessId -eq 0 -or $ownerPid -eq $ProcessId) {
            $winCond = New-Object System.Windows.Automation.PropertyCondition(
                [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
                [System.Windows.Automation.ControlType]::Window)

            foreach ($w in $child.FindAll([System.Windows.Automation.TreeScope]::Descendants, $winCond)) {
                if (Select-Window -Element $w) { break }
            }
        }

        $child = $walker.GetNextSibling($child)
    }
}

if ($handle -eq [IntPtr]::Zero) {
    Write-Output "Окно не найдено (PID=$ProcessId, заголовок '$TitleLike')"
    exit 1
}

[void][NativeMethods]::ShowWindow($handle, 9)

# При заданных размерах окно вписывается в рабочую область экрана,
# чтобы снимок не перекрывался панелью задач.
if ($Width -gt 0 -and $Height -gt 0) {
    [void][NativeMethods]::MoveWindow($handle, 20, 20, $Width, $Height, $true)
    Start-Sleep -Milliseconds 700
}

[void][NativeMethods]::SetForegroundWindow($handle)
Start-Sleep -Milliseconds 900

$rect = New-Object NativeMethods+RECT
[void][NativeMethods]::GetWindowRect($handle, [ref]$rect)

$width = $rect.Right - $rect.Left
$height = $rect.Bottom - $rect.Top

if ($width -le 0 -or $height -le 0) {
    Write-Output "Некорректный размер окна: ${width}x${height}"
    exit 1
}

$bitmap = New-Object System.Drawing.Bitmap($width, $height)
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$graphics.CopyFromScreen($rect.Left, $rect.Top, 0, 0, $bitmap.Size)

$full = [System.IO.Path]::GetFullPath($Output)
$dir = [System.IO.Path]::GetDirectoryName($full)

if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }

$bitmap.Save($full, [System.Drawing.Imaging.ImageFormat]::Png)

$graphics.Dispose()
$bitmap.Dispose()

Write-Output "Снимок '$windowName': $full ($width x $height)"
