# -*- coding: utf-8 -*-
# Демонстрационные снимки форм для презентации к защите.
#
# Зачем отдельный скрипт: обычный Capture-Deck.ps1 снимает формы под
# администратором, из-за чего корзина получается пустой, а в истории
# заказов нет покупательских записей. Здесь снимаются три снимка
# в правдоподобном состоянии:
#
#   07_cart.png             — корзина покупателя с товарами и итогом;
#   06_my_orders.png        — история заказов покупателя (maria);
#   04_manager_orders.png   — обработка заказов сотрудником (ivan).
#
# Снимок делается через PrintWindow: в кадр не попадают панель задач
# и соседние окна. Координаты элементов пишутся в callouts.json.
#
# Запуск:
#   powershell -File tools/uia/Capture-Demo.ps1

param(
    [int]$AppProcessId = 0,
    [string]$OutputDir = 'docs\screens'
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

Add-Type @'
using System;
using System.Runtime.InteropServices;

public static class DemoNative {
    [DllImport("user32.dll")]
    public static extern bool PrintWindow(IntPtr hWnd, IntPtr hdcBlt, uint nFlags);

    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT r);

    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left, Top, Right, Bottom; }
}
'@

$root = [System.Windows.Automation.AutomationElement]::RootElement
$walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
$scope = [System.Windows.Automation.TreeScope]::Descendants

$script:rects = @{}

function Get-TopWindows {
    $list = New-Object System.Collections.ArrayList
    $child = $walker.GetFirstChild($root)

    while ($null -ne $child) {
        [void]$list.Add($child)
        $child = $walker.GetNextSibling($child)
    }

    return $list
}

function Find-Window {
    param([string]$NameLike = "", [int]$OwnerProcessId = 0)

    foreach ($w in Get-TopWindows) {
        try {
            $name = $w.Current.Name
            $ownerPid = $w.Current.ProcessId
        } catch { continue }

        if ($OwnerProcessId -ne 0 -and $ownerPid -ne $OwnerProcessId) { continue }
        if ($NameLike -ne "" -and $name -notlike "*$NameLike*") { continue }

        return $w
    }

    if ($NameLike -eq "") { return $null }

    $cond = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
        [System.Windows.Automation.ControlType]::Window)

    foreach ($top in Get-TopWindows) {
        try {
            if ($OwnerProcessId -ne 0 -and $top.Current.ProcessId -ne $OwnerProcessId) { continue }
        } catch { continue }

        foreach ($w in $top.FindAll($scope, $cond)) {
            try {
                if ($w.Current.Name -like "*$NameLike*") { return $w }
            } catch { }
        }
    }

    return $null
}

function Find-ChildWindow {
    param([System.Windows.Automation.AutomationElement]$Parent, [string]$NameLike)

    $cond = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
        [System.Windows.Automation.ControlType]::Window)

    foreach ($w in $Parent.FindAll($scope, $cond)) {
        try {
            if ($w.Current.Name -like "*$NameLike*") { return $w }
        } catch { }
    }

    return $null
}

function Find-ById {
    param([System.Windows.Automation.AutomationElement]$Parent, [string]$AutomationId)

    $cond = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::AutomationIdProperty, $AutomationId)

    return $Parent.FindFirst($scope, $cond)
}

function Find-AllByName {
    param([System.Windows.Automation.AutomationElement]$Parent, [string]$Name, [string]$Type = "")

    $all = $Parent.FindAll($scope, [System.Windows.Automation.Condition]::TrueCondition)
    $result = New-Object System.Collections.ArrayList

    foreach ($element in $all) {
        try {
            if ($element.Current.Name -ne $Name) { continue }

            if ($Type -ne "") {
                $controlType = $element.Current.ControlType.ProgrammaticName -replace 'ControlType\.', ''

                if ($controlType -ne $Type) { continue }
            }

            [void]$result.Add($element)
        } catch { }
    }

    return $result
}

function Find-AllByType {
    param([System.Windows.Automation.AutomationElement]$Parent, [string]$Type)

    $cond = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
        [System.Windows.Automation.ControlType]::$Type)

    return $Parent.FindAll($scope, $cond)
}

function Invoke-Element {
    param([System.Windows.Automation.AutomationElement]$Element)

    $pattern = $Element.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
    $pattern.Invoke()
}

function Close-Window {
    param([System.Windows.Automation.AutomationElement]$Window)

    try {
        $pattern = $Window.GetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern)
        $pattern.Close()
    } catch { }

    Start-Sleep -Milliseconds 900
}

function Set-Text {
    param([System.Windows.Automation.AutomationElement]$Element, [string]$Text)

    $pattern = $null

    if ($Element.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern, [ref]$pattern)) {
        $pattern.SetValue($Text)
        return
    }

    throw "Поле ввода не поддерживает шаблон значения."
}

# Снимок окна через PrintWindow: содержимое окна без посторонних окон.
function Save-WindowShot {
    param(
        [System.Windows.Automation.AutomationElement]$Window,
        [string]$Path
    )

    $handle = [IntPtr]$Window.Current.NativeWindowHandle

    if ($handle -eq [IntPtr]::Zero) {
        Write-Host "  ! окно без дескриптора — снимок не сделан"
        return $null
    }

    $native = New-Object DemoNative+RECT
    [void][DemoNative]::GetWindowRect($handle, [ref]$native)

    $w = $native.Right - $native.Left
    $h = $native.Bottom - $native.Top

    if ($w -le 0 -or $h -le 0) {
        $box = $Window.Current.BoundingRectangle
        $w = [int][math]::Round($box.Width)
        $h = [int][math]::Round($box.Height)
    }

    [void][DemoNative]::ShowWindow($handle, 9)
    [void][DemoNative]::SetForegroundWindow($handle)
    Start-Sleep -Milliseconds 700

    $bitmap = New-Object System.Drawing.Bitmap($w, $h)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $hdc = $graphics.GetHdc()

    # PW_RENDERFULLCONTENT = 2 — рисует и содержимое окна, и DWM-эффекты
    [void][DemoNative]::PrintWindow($handle, $hdc, 2)

    $graphics.ReleaseHdc($hdc)
    $graphics.Dispose()

    $full = [System.IO.Path]::GetFullPath($Path)
    $dir = [System.IO.Path]::GetDirectoryName($full)

    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }

    $bitmap.Save($full, [System.Drawing.Imaging.ImageFormat]::Png)
    $bitmap.Dispose()

    Write-Host "  снимок: $full ($w x $h)"

    return [pscustomobject]@{ X = $native.Left; Y = $native.Top; W = $w; H = $h }
}

# Запоминает прямоугольник элемента в координатах снимка окна.
function Save-Rect {
    param(
        [System.Windows.Automation.AutomationElement]$Element,
        [string]$Key,
        $WindowBox
    )

    if ($null -eq $Element -or $null -eq $WindowBox) { return }

    try { $box = $Element.Current.BoundingRectangle } catch { return }

    if ($box.Width -le 0 -or $box.Height -le 0) { return }

    $script:rects[$Key] = @{
        x = [int][math]::Round($box.Left - $WindowBox.X)
        y = [int][math]::Round($box.Top - $WindowBox.Y)
        w = [int][math]::Round($box.Width)
        h = [int][math]::Round($box.Height)
    }
}

# --- Вход и выход из приложения ---

function Get-MainWindow {
    param([int]$OwnerProcessId)

    return Find-Window -NameLike 'Магазин автозапчастей' -OwnerProcessId $OwnerProcessId
}

function Get-CurrentUserName {
    param([System.Windows.Automation.AutomationElement]$Main)

    $element = Find-ById -Parent $Main -AutomationId 'CurrentUserText'
    if ($null -eq $element) { return '' }

    return $element.Current.Name
}

function Login-User {
    param(
        [System.Windows.Automation.AutomationElement]$Main,
        [string]$Login,
        [string]$Password
    )

    $account = Find-ById -Parent $Main -AutomationId 'AccountButton'

    if ($null -eq $account) {
        throw "Кнопка «Вход / Регистрация» не найдена."
    }

    Invoke-Element -Element $account
    Start-Sleep -Seconds 2

    $auth = Find-ChildWindow -Parent $Main -NameLike 'Авторизация'

    if ($null -eq $auth) {
        throw "Окно авторизации не открылось."
    }

    $tab = Find-AllByName -Parent $auth -Name 'Вход' -Type 'TabItem'
    if ($tab.Count -gt 0) {
        try {
            $selection = $tab[0].GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern)
            $selection.Select()
            Start-Sleep -Milliseconds 500
        } catch { }
    }

    $edits = Find-AllByType -Parent $auth -Type 'Edit'

    if ($edits.Count -lt 2) {
        throw "Не найдены поля логина и пароля."
    }

    Set-Text -Element $edits[0] -Text $Login
    Start-Sleep -Milliseconds 300

    $auth = Find-ChildWindow -Parent $Main -NameLike 'Авторизация'
    $edits = Find-AllByType -Parent $auth -Type 'Edit'
    Set-Text -Element $edits[1] -Text $Password
    Start-Sleep -Milliseconds 300

    $auth = Find-ChildWindow -Parent $Main -NameLike 'Авторизация'
    $loginButton = Find-AllByName -Parent $auth -Name 'Войти' -Type 'Button'

    if ($loginButton.Count -eq 0) {
        throw "Кнопка «Войти» не найдена."
    }

    Invoke-Element -Element $loginButton[0]
    Start-Sleep -Seconds 3

    $Main = Get-MainWindow -OwnerProcessId $AppProcessId
    $name = Get-CurrentUserName -Main $Main

    Write-Host "  вход: $Login -> $name"

    return $Main
}

function Logout-User {
    param([System.Windows.Automation.AutomationElement]$Main)

    $logout = Find-ById -Parent $Main -AutomationId 'LogoutButton'

    if ($null -eq $logout) {
        Write-Host "  ! кнопка «Выход» не найдена"
        return
    }

    Invoke-Element -Element $logout
    Start-Sleep -Seconds 2
}

# --- Подготовка ---

$projectRoot = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSCommandPath))
$out = if ([System.IO.Path]::IsPathRooted($OutputDir)) { $OutputDir } else { Join-Path $projectRoot $OutputDir }

if ($AppProcessId -eq 0) {
    $exe = Get-ChildItem (Join-Path $projectRoot 'scr\Ktpt.CourseOrder\bin') -Recurse -Filter 'Ktpt.CourseOrder.exe' -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1

    if ($null -eq $exe) {
        Write-Output "Собранный исполняемый файл приложения не найден."
        exit 1
    }

    Write-Output "Запуск приложения: $($exe.FullName)"

    $proc = Start-Process -FilePath $exe.FullName -WorkingDirectory $exe.DirectoryName -PassThru
    $AppProcessId = $proc.Id

    # Окно появляется не сразу: сначала создаётся база и загружаются данные.
    for ($i = 0; $i -lt 40; $i++) {
        Start-Sleep -Milliseconds 500

        $main = Get-MainWindow -OwnerProcessId $AppProcessId
        if ($null -ne $main) { break }
    }
}

$main = Get-MainWindow -OwnerProcessId $AppProcessId

if ($null -eq $main) {
    Write-Output "Главное окно не найдено (PID=$AppProcessId)."
    exit 1
}

Write-Output "Демонстрационные снимки, PID=$AppProcessId"
Write-Output ""

# --- 1. Корзина покупателя с товарами ---

Write-Output "=== 1. Корзина ==="

if ((Get-CurrentUserName -Main $main) -eq 'Гость') {
    $main = Login-User -Main $main -Login 'ander' -Password '8942'
}

# Два разных товара в корзине.
$addButtons = Find-AllByName -Parent $main -Name 'Добавить в корзину' -Type 'Button'
Write-Host "  кнопок «Добавить в корзину»: $($addButtons.Count)"

foreach ($button in @($addButtons[0], $addButtons[1])) {
    if ($null -eq $button) { continue }
    Invoke-Element -Element $button
    Start-Sleep -Milliseconds 700
}

$cart = Find-ById -Parent $main -AutomationId 'CartButton'
Invoke-Element -Element $cart
Start-Sleep -Seconds 2

$cartWindow = Find-Window -NameLike 'Корзина' -OwnerProcessId $AppProcessId
$box = Save-WindowShot -Window $cartWindow -Path (Join-Path $out '07_cart.png')

if ($box) {
    foreach ($e in $cartWindow.FindAll($scope, [System.Windows.Automation.Condition]::TrueCondition)) {
        try {
            $type = $e.Current.ControlType.ProgrammaticName -replace 'ControlType\.', ''
            $name = $e.Current.Name

            if ($type -eq 'Text' -and $name -like 'Итого:*') { Save-Rect -Element $e -Key 'cart.total' -WindowBox $box }
            if ($type -eq 'Button' -and $name -eq 'Оплатить') { Save-Rect -Element $e -Key 'cart.pay' -WindowBox $box }
        } catch { }
    }
}

Close-Window -Window $cartWindow

# --- 2. История заказов покупателя ---

Write-Output "=== 2. Мои заказы (покупатель) ==="

Logout-User -Main $main
$main = Get-MainWindow -OwnerProcessId $AppProcessId
$main = Login-User -Main $main -Login 'maria' -Password '1234'

$myOrders = Find-ById -Parent $main -AutomationId 'MyOrdersButton'
Invoke-Element -Element $myOrders
Start-Sleep -Seconds 2

$ordersWindow = Find-Window -NameLike 'Мои заказы' -OwnerProcessId $AppProcessId
$box = Save-WindowShot -Window $ordersWindow -Path (Join-Path $out '06_my_orders.png')

if ($box) {
    foreach ($e in $ordersWindow.FindAll($scope, [System.Windows.Automation.Condition]::TrueCondition)) {
        try {
            if ($e.Current.ControlType.ProgrammaticName -match 'Text' -and $e.Current.Name -like 'Всего заказов:*') {
                Save-Rect -Element $e -Key 'orders.total' -WindowBox $box
            }
        } catch { }
    }
}

Close-Window -Window $ordersWindow

# --- 3. Обработка заказов сотрудником ---

Write-Output "=== 3. Обработка заказов (сотрудник) ==="

Logout-User -Main $main
$main = Get-MainWindow -OwnerProcessId $AppProcessId
$main = Login-User -Main $main -Login 'ivan' -Password '5678'

$manager = Find-ById -Parent $main -AutomationId 'ManagerButton'

if ($null -ne $manager) {
    Invoke-Element -Element $manager
    Start-Sleep -Seconds 2

    $managerWindow = Find-Window -NameLike 'Обработка заказов' -OwnerProcessId $AppProcessId
    $box = Save-WindowShot -Window $managerWindow -Path (Join-Path $out '04_manager_orders.png')

    if ($box) {
        foreach ($e in $managerWindow.FindAll($scope, [System.Windows.Automation.Condition]::TrueCondition)) {
            try {
                $type = $e.Current.ControlType.ProgrammaticName -replace 'ControlType\.', ''
                $name = $e.Current.Name

                if ($type -eq 'ComboBox') { Save-Rect -Element $e -Key 'manager.statusCombo' -WindowBox $box }

                # Кнопка «Отменить» — берём первую (верхний заказ списка),
                # нижние строки окна обрезаны и не годятся для выноски.
                if ($type -eq 'Button' -and $name -like 'Отменить*' -and -not $script:rects.ContainsKey('manager.cancel')) {
                    Save-Rect -Element $e -Key 'manager.cancel' -WindowBox $box
                }
            } catch { }
        }
    }

    Close-Window -Window $managerWindow
}

# --- Сохранение координат выносок ---

$calloutFile = Join-Path $out 'callouts.json'
$all = @{}

if (Test-Path -LiteralPath $calloutFile) {
    $loaded = Get-Content -LiteralPath $calloutFile -Raw -Encoding UTF8 | ConvertFrom-Json

    foreach ($property in $loaded.PSObject.Properties) {
        $all[$property.Name] = $property.Value
    }
}

foreach ($key in $script:rects.Keys) {
    $all[$key] = $script:rects[$key]
}

$json = $all | ConvertTo-Json -Depth 5
[System.IO.File]::WriteAllText($calloutFile, $json, (New-Object System.Text.UTF8Encoding($true)))

Write-Output ""
Write-Output "Координаты выносок обновлены: $calloutFile"
Write-Output "Обновлено ключей: $($script:rects.Keys.Count)"
exit 0
