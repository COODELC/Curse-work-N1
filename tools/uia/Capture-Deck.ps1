# -*- coding: utf-8 -*-
# Снимки экранных форм для презентации.
#
# Отличия от Capture-Forms.ps1:
#   * снимок делается через PrintWindow, поэтому в кадр не попадают
#     панель задач и соседние окна;
#   * запоминаются координаты элементов управления в пикселях снимка —
#     они используются для расстановки выносок на слайдах;
#   * все формы снимаются в натуральном разрешении, без масштабирования.
#
# Запуск:
#   powershell -File tools/uia/Capture-Deck.ps1 -ProcessId 1234

param(
    [int]$ProcessId = 0,
    [string]$OutputDir = 'docs\screens'
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

Add-Type @'
using System;
using System.Runtime.InteropServices;

public static class DeckNative {
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

    # Сначала верхний уровень, затем — потомки найденного окна:
    # модальные диалоги Avalonia входят в дерево UIA как вложенные окна.
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

    # Границы окна берём из Win32: они не зависят от того, видимо ли окно.
    $native = New-Object DeckNative+RECT
    [void][DeckNative]::GetWindowRect($handle, [ref]$native)

    $w = $native.Right - $native.Left
    $h = $native.Bottom - $native.Top

    if ($w -le 0 -or $h -le 0) {
        $box = $Window.Current.BoundingRectangle
        $w = [int][math]::Round($box.Width)
        $h = [int][math]::Round($box.Height)
    }

    if ($w -le 0 -or $h -le 0) {
        Write-Host "  ! некорректный размер окна"
        return $null
    }

    [void][DeckNative]::ShowWindow($handle, 9)
    [void][DeckNative]::SetForegroundWindow($handle)
    Start-Sleep -Milliseconds 700

    $bitmap = New-Object System.Drawing.Bitmap($w, $h)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $hdc = $graphics.GetHdc()

    # PW_RENDERFULLCONTENT = 2 — рисует и содержимое окна, и DWM-эффекты
    [void][DeckNative]::PrintWindow($handle, $hdc, 2)

    $graphics.ReleaseHdc($hdc)
    $graphics.Dispose()

    $full = [System.IO.Path]::GetFullPath($Path)
    $dir = [System.IO.Path]::GetDirectoryName($full)

    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }

    $bitmap.Save($full, [System.Drawing.Imaging.ImageFormat]::Png)
    $bitmap.Dispose()

    Write-Host "  снимок: $full ($w x $h)"

    # Координаты отсчитываются от левого верхнего угла снимка,
    # то есть от рамки окна, а не от его клиентской области.
    return [pscustomobject]@{
        X = $native.Left
        Y = $native.Top
        W = $w
        H = $h
    }
}

# Запоминает прямоугольник элемента в координатах снимка окна.
function Save-Rect {
    param(
        [System.Windows.Automation.AutomationElement]$Element,
        [string]$Key,
        $WindowBox
    )

    if ($null -eq $Element -or $null -eq $WindowBox) { return }

    try {
        $box = $Element.Current.BoundingRectangle
    } catch { return }

    if ($box.Width -le 0 -or $box.Height -le 0) { return }

    $script:rects[$Key] = @{
        x = [int][math]::Round($box.Left - $WindowBox.X)
        y = [int][math]::Round($box.Top - $WindowBox.Y)
        w = [int][math]::Round($box.Width)
        h = [int][math]::Round($box.Height)
    }
}

# --- Подготовка ---

if ($ProcessId -eq 0) {
    $proc = Get-Process Ktpt.CourseOrder -ErrorAction SilentlyContinue | Select-Object -First 1

    if ($null -eq $proc) {
        Write-Output "Приложение не запущено."
        exit 1
    }

    $ProcessId = $proc.Id
}

$out = if ([System.IO.Path]::IsPathRooted($OutputDir)) { $OutputDir } else { Join-Path (Get-Location).Path $OutputDir }

Write-Output "Снимки экранных форм, PID=$ProcessId"
Write-Output ""

# Окна, оставшиеся от прошлого прогона, блокируют главное окно — закрываем их.
$winCond = New-Object System.Windows.Automation.PropertyCondition(
    [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
    [System.Windows.Automation.ControlType]::Window)

foreach ($top in Get-TopWindows) {
    try {
        if ($top.Current.ProcessId -ne $ProcessId) { continue }
    } catch { continue }

    foreach ($w in $top.FindAll($scope, $winCond)) {
        try {
            if ($w.Current.Name -like '*Магазин автозапчастей*') { continue }
            Close-Window -Window $w
        } catch { }
    }

    try {
        if ($top.Current.Name -notlike '*Магазин автозапчастей*') { Close-Window -Window $top }
    } catch { }
}

Start-Sleep -Milliseconds 500

$main = Find-Window -NameLike 'Магазин автозапчастей' -OwnerProcessId $ProcessId

if ($null -eq $main) {
    Write-Output "Главное окно не найдено."
    exit 1
}

# --- 1. Главное окно ---

Write-Output "=== 1. Главное окно ==="

$main = Find-Window -NameLike 'Магазин автозапчастей' -OwnerProcessId $ProcessId
$box = Save-WindowShot -Window $main -Path (Join-Path $out '01_main.png')

if ($box) {
    $script:rects['main.header'] = @{ x = 44; y = 50; w = 1072; h = 104 }
    Save-Rect -Element (Find-ById -Parent $main -AutomationId 'CartButton') -Key 'main.cartButton' -WindowBox $box
    Save-Rect -Element (Find-ById -Parent $main -AutomationId 'MyOrdersButton') -Key 'main.myOrders' -WindowBox $box
    Save-Rect -Element (Find-ById -Parent $main -AutomationId 'AdminButton') -Key 'main.admin' -WindowBox $box
    Save-Rect -Element (Find-ById -Parent $main -AutomationId 'CurrentUserText') -Key 'main.user' -WindowBox $box

    $all = $main.FindAll($scope, [System.Windows.Automation.Condition]::TrueCondition)

    # Поле поиска: сначала по AutomationId, затем по подсказке поля.
    Save-Rect -Element (Find-ById -Parent $main -AutomationId 'SearchBox') -Key 'main.search' -WindowBox $box

    foreach ($e in $all) {
        try {
            $type = $e.Current.ControlType.ProgrammaticName -replace 'ControlType\.', ''

            if ($type -eq 'Edit' -and $e.Current.Name -like '*Поиск*') {
                Save-Rect -Element $e -Key 'main.search' -WindowBox $box
            }
            if ($type -eq 'ComboBox') {
                Save-Rect -Element $e -Key 'main.category' -WindowBox $box
            }
        } catch { }
    }

    $firstAdd = (Find-AllByName -Parent $main -Name 'Добавить в корзину' -Type 'Button')
    if ($firstAdd.Count -gt 0) { Save-Rect -Element $firstAdd[0] -Key 'main.addToCart' -WindowBox $box }
}

# --- 2. Описание товара ---

Write-Output "=== 2. Описание товара ==="

$details = (Find-AllByName -Parent $main -Name 'Описание товара' -Type 'Button')
if ($details.Count -gt 0) {
    Invoke-Element -Element $details[0]
    Start-Sleep -Seconds 2

    $w = Find-Window -NameLike 'О товаре' -OwnerProcessId $ProcessId
    $box = Save-WindowShot -Window $w -Path (Join-Path $out '02_product_details.png')

    if ($box) {
        $all = $w.FindAll($scope, [System.Windows.Automation.Condition]::TrueCondition)

        foreach ($e in $all) {
            try {
                $type = $e.Current.ControlType.ProgrammaticName -replace 'ControlType\.', ''
                $name = $e.Current.Name

                if ($type -eq 'Text' -and $name -like 'Информация о товаре*') { Save-Rect -Element $e -Key 'details.heading' -WindowBox $box }
                if ($type -eq 'Text' -and $name -like 'Описание*') { Save-Rect -Element $e -Key 'details.description' -WindowBox $box }
                if ($type -eq 'Text' -and $name -like 'Артикул*') { Save-Rect -Element $e -Key 'details.articule' -WindowBox $box }
            } catch { }
        }
    }

    Close-Window -Window $w
}

# --- 3. Отзывы ---

Write-Output "=== 3. Отзывы о товаре ==="

$main = Find-Window -NameLike 'Магазин автозапчастей' -OwnerProcessId $ProcessId

# Ссылка на отзывы содержит число, поэтому ищем по началу названия.
$reviews = $null
foreach ($e in $main.FindAll($scope, [System.Windows.Automation.Condition]::TrueCondition)) {
    try {
        if ($e.Current.ControlType.ProgrammaticName -match 'Button' -and $e.Current.Name -like 'Отзывов:*') {
            $reviews = $e
            break
        }
    } catch { }
}

if ($null -ne $reviews) {
    Invoke-Element -Element $reviews
    Start-Sleep -Seconds 2

    $w = Find-Window -NameLike 'Отзывы о товаре' -OwnerProcessId $ProcessId
    Save-WindowShot -Window $w -Path (Join-Path $out '03_reviews.png') | Out-Null
    Close-Window -Window $w
}

# --- 4. Обработка заказов ---

Write-Output "=== 4. Обработка заказов ==="

$main = Find-Window -NameLike 'Магазин автозапчастей' -OwnerProcessId $ProcessId
$manager = Find-ById -Parent $main -AutomationId 'ManagerButton'

if ($null -ne $manager) {
    Invoke-Element -Element $manager
    Start-Sleep -Seconds 2

    $w = Find-Window -NameLike 'Обработка заказов' -OwnerProcessId $ProcessId
    $box = Save-WindowShot -Window $w -Path (Join-Path $out '04_manager_orders.png')

    if ($box) {
        $all = $w.FindAll($scope, [System.Windows.Automation.Condition]::TrueCondition)

        foreach ($e in $all) {
            try {
                $type = $e.Current.ControlType.ProgrammaticName -replace 'ControlType\.', ''

                if ($type -eq 'ComboBox') { Save-Rect -Element $e -Key 'manager.statusCombo' -WindowBox $box }
                if ($type -eq 'Button' -and $e.Current.Name -like 'Сменить статус*') { Save-Rect -Element $e -Key 'manager.setStatus' -WindowBox $box }
                if ($type -eq 'Button' -and $e.Current.Name -like 'Отменить*') { Save-Rect -Element $e -Key 'manager.cancel' -WindowBox $box }
            } catch { }
        }
    }

    Close-Window -Window $w
}

# --- 5. Управление магазином ---

Write-Output "=== 5. Управление магазином ==="

$main = Find-Window -NameLike 'Магазин автозапчастей' -OwnerProcessId $ProcessId
$admin = Find-ById -Parent $main -AutomationId 'AdminButton'

if ($null -ne $admin) {
    Invoke-Element -Element $admin
    Start-Sleep -Seconds 2

    $w = Find-Window -NameLike 'Управление магазином' -OwnerProcessId $ProcessId
    Save-WindowShot -Window $w -Path (Join-Path $out '05_admin.png') | Out-Null
    Close-Window -Window $w
}

# --- 6. Мои заказы ---

Write-Output "=== 6. Мои заказы ==="

$main = Find-Window -NameLike 'Магазин автозапчастей' -OwnerProcessId $ProcessId
$myOrders = Find-ById -Parent $main -AutomationId 'MyOrdersButton'

if ($null -ne $myOrders) {
    Invoke-Element -Element $myOrders
    Start-Sleep -Seconds 2

    $w = Find-Window -NameLike 'Мои заказы' -OwnerProcessId $ProcessId
    $box = Save-WindowShot -Window $w -Path (Join-Path $out '06_my_orders.png')

    if ($box) {
        $all = $w.FindAll($scope, [System.Windows.Automation.Condition]::TrueCondition)

        foreach ($e in $all) {
            try {
                if ($e.Current.ControlType.ProgrammaticName -match 'Text' -and $e.Current.Name -like 'Всего заказов:*') {
                    Save-Rect -Element $e -Key 'orders.total' -WindowBox $box
                }
            } catch { }
        }
    }

    Close-Window -Window $w
}

# --- 7. Корзина ---

Write-Output "=== 7. Корзина ==="

$main = Find-Window -NameLike 'Магазин автозапчастей' -OwnerProcessId $ProcessId
$cart = Find-ById -Parent $main -AutomationId 'CartButton'

if ($null -ne $cart) {
    Invoke-Element -Element $cart
    Start-Sleep -Seconds 2

    $w = Find-Window -NameLike 'Корзина' -OwnerProcessId $ProcessId
    $box = Save-WindowShot -Window $w -Path (Join-Path $out '07_cart.png')

    if ($box) {
        $all = $w.FindAll($scope, [System.Windows.Automation.Condition]::TrueCondition)

        foreach ($e in $all) {
            try {
                $type = $e.Current.ControlType.ProgrammaticName -replace 'ControlType\.', ''
                $name = $e.Current.Name

                if ($type -eq 'Text' -and $name -like 'Итого:*') { Save-Rect -Element $e -Key 'cart.total' -WindowBox $box }
                if ($type -eq 'Button' -and $name -eq 'Оплатить') { Save-Rect -Element $e -Key 'cart.pay' -WindowBox $box }
            } catch { }
        }
    }

    Close-Window -Window $w
}

# --- 8. Профиль ---

Write-Output "=== 8. Профиль ==="

$main = Find-Window -NameLike 'Магазин автозапчастей' -OwnerProcessId $ProcessId
$profile = Find-ById -Parent $main -AutomationId 'ProfileButton'

if ($null -ne $profile) {
    Invoke-Element -Element $profile
    Start-Sleep -Seconds 2

    $w = Find-Window -NameLike 'Профиль' -OwnerProcessId $ProcessId
    Save-WindowShot -Window $w -Path (Join-Path $out '08_profile.png') | Out-Null
    Close-Window -Window $w
}

# --- 9. Управление ролями ---

Write-Output "=== 9. Управление ролями ==="

$main = Find-Window -NameLike 'Магазин автозапчастей' -OwnerProcessId $ProcessId
$roles = Find-ById -Parent $main -AutomationId 'RolesButton'

if ($null -ne $roles) {
    Invoke-Element -Element $roles
    Start-Sleep -Seconds 2

    $w = Find-Window -NameLike 'Управление ролями' -OwnerProcessId $ProcessId
    Save-WindowShot -Window $w -Path (Join-Path $out '09_roles.png') | Out-Null
    Close-Window -Window $w
}

# --- 10. Гостевой режим ---

Write-Output "=== 10. Гостевой режим ==="

$main = Find-Window -NameLike 'Магазин автозапчастей' -OwnerProcessId $ProcessId
$logout = Find-ById -Parent $main -AutomationId 'LogoutButton'

if ($null -ne $logout) {
    Invoke-Element -Element $logout
    Start-Sleep -Seconds 2

    $main = Find-Window -NameLike 'Магазин автозапчастей' -OwnerProcessId $ProcessId
    $box = Save-WindowShot -Window $main -Path (Join-Path $out '10_guest.png')

    if ($box) {
        $all = $main.FindAll($scope, [System.Windows.Automation.Condition]::TrueCondition)

        foreach ($e in $all) {
            try {
                $type = $e.Current.ControlType.ProgrammaticName -replace 'ControlType\.', ''

                if ($type -eq 'Edit' -and $e.Current.Name -like '*Поиск*') {
                    Save-Rect -Element $e -Key 'guest.search' -WindowBox $box
                }
                if ($type -eq 'Button' -and $e.Current.Name -like '*Вход*') {
                    Save-Rect -Element $e -Key 'guest.login' -WindowBox $box
                }
            } catch { }
        }
    }
}

# --- Сохранение координат ---

$jsonPath = Join-Path $out 'callouts.json'
$script:rects | ConvertTo-Json -Depth 4 | Set-Content -Path $jsonPath -Encoding UTF8

Write-Output ""
Write-Output "Координаты элементов: $jsonPath"
Write-Output "Записано прямоугольников: $($script:rects.Count)"
