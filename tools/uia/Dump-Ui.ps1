# Скрипты UI Automation для приложения «Магазин автозапчастей» (Avalonia, Windows).
# Использование:
#   powershell -File tools/uia/Dump-Ui.ps1 -WindowTitle "Магазин автозапчастей"

param(
    [string]$WindowTitle = "",
    [int]$ProcessId = 0,
    [string]$RootName = ""
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

$root = [System.Windows.Automation.AutomationElement]::RootElement

$walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker

function Get-TopWindows {
    param([System.Windows.Automation.AutomationElement]$Root)

    $result = New-Object System.Collections.ArrayList
    $child = $walker.GetFirstChild($Root)

    while ($null -ne $child) {
        [void]$result.Add($child)
        $child = $walker.GetNextSibling($child)
    }

    return $result
}

function Find-AppWindow {
    $windows = Get-TopWindows -Root $root

    foreach ($w in $windows) {
        try {
            $name = $w.Current.Name
            $windowProcessId = $w.Current.ProcessId
        } catch {
            continue
        }

        if ($ProcessId -ne 0 -and $windowProcessId -ne $ProcessId) { continue }
        if ($WindowTitle -ne "" -and $name -notlike "*$WindowTitle*") { continue }

        return $w
    }

    return $null
}

function Dump-Tree {
    param(
        [System.Windows.Automation.AutomationElement]$Element,
        [int]$Depth = 0,
        [int]$MaxDepth = 6
    )

    if ($null -eq $Element -or $Depth -gt $MaxDepth) { return }

    $c = $Element.Current
    $name = $c.Name
    $type = $c.ControlType.ProgrammaticName -replace 'ControlType\.', ''
    $autoId = $c.AutomationId

    $patterns = @()
    foreach ($p in @(
        [System.Windows.Automation.InvokePattern]::Pattern,
        [System.Windows.Automation.ValuePattern]::Pattern,
        [System.Windows.Automation.SelectionItemPattern]::Pattern,
        [System.Windows.Automation.TogglePattern]::Pattern,
        [System.Windows.Automation.ExpandCollapsePattern]::Pattern
    )) {
        if ($Element.TryGetCurrentPattern($p, [ref]$null)) { $patterns += $p.ProgrammaticName -replace 'Pattern\.', '' }
    }

    $line = ('  ' * $Depth) + "[$type] name='$name'"
    if ($autoId) { $line += " id='$autoId'" }
    if ($patterns.Count -gt 0) { $line += " patterns=" + ($patterns -join ',') }

    Write-Output $line

    $child = $walker.GetFirstChild($Element)

    while ($null -ne $child) {
        Dump-Tree -Element $child -Depth ($Depth + 1) -MaxDepth $MaxDepth
        $child = $walker.GetNextSibling($child)
    }
}

$window = Find-AppWindow

if ($null -eq $window) {
    Write-Output "ОКНО НЕ НАЙДЕНО"
    exit 1
}

Write-Output "ОКНО: '$($window.Current.Name)' PID=$($window.Current.ProcessId)"
Write-Output ""

if ($RootName -ne "") {
    $cond = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::NameProperty, $RootName)

    $target = $window.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $cond)

    if ($null -eq $target) {
        Write-Output "ЭЛЕМЕНТ '$RootName' НЕ НАЙДЕН"
        exit 1
    }

    Write-Output "--- Корень: $RootName ---"
    Dump-Tree -Element $target
} else {
    Dump-Tree -Element $window
}
