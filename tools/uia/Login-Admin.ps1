# Вход в приложение «Магазин автозапчастей» через UI Automation.
#   powershell -File tools/uia/Login-Admin.ps1 -Login ander -Password 8942

param(
    [string]$Login = "ander",
    [string]$Password = "8942",
    [int]$AppProcessId = 0,
    [switch]$DumpOnly
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Windows.Forms

Add-Type @'
using System;
using System.Runtime.InteropServices;

public static class NativeMethods {
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    public static extern bool SetCursorPos(int x, int y);

    [DllImport("user32.dll")]
    public static extern void mouse_event(uint flags, uint dx, uint dy, uint data, UIntPtr extra);

    [DllImport("user32.dll")]
    public static extern bool BringWindowToTop(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();
}
'@

$root = [System.Windows.Automation.AutomationElement]::RootElement
$walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
$scope = [System.Windows.Automation.TreeScope]::Descendants

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

    return $null
}

function Find-ByName {
    param([System.Windows.Automation.AutomationElement]$Parent, [string]$Name, [string]$Type = "")

    $conditions = New-Object System.Collections.ArrayList
    [void]$conditions.Add((New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::NameProperty, $Name)))

    if ($Type -ne "") {
        [void]$conditions.Add((New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
            [System.Windows.Automation.ControlType]::$Type)))
    }

    if ($conditions.Count -eq 1) {
        $cond = $conditions[0]
    } else {
        $cond = New-Object System.Windows.Automation.AndCondition(
            $conditions[0], $conditions[1])
    }

    return $Parent.FindFirst($scope, $cond)
}

function Set-Foreground {
    param([System.Windows.Automation.AutomationElement]$Element)

    try {
        $handle = [IntPtr]$Element.Current.NativeWindowHandle

        if ($handle -ne [IntPtr]::Zero) {
            [void][NativeMethods]::ShowWindow($handle, 9)
            [void][NativeMethods]::SetForegroundWindow($handle)
            Start-Sleep -Milliseconds 400
        }
    } catch {
        # Если окно не удалось активировать, ввод с клавиатуры может не сработать.
    }
}

function Find-ChildWindow {
    param([System.Windows.Automation.AutomationElement]$Parent, [string]$NameLike)

    $cond = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
        [System.Windows.Automation.ControlType]::Window)

    $windows = $Parent.FindAll($scope, $cond)

    foreach ($w in $windows) {
        try {
            if ($w.Current.Name -like "*$NameLike*") { return $w }
        } catch { }
    }

    return $null
}

function Find-ById {    param([System.Windows.Automation.AutomationElement]$Parent, [string]$AutomationId)

    $cond = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::AutomationIdProperty, $AutomationId)

    return $Parent.FindFirst($scope, $cond)
}

function Find-All {
    param([System.Windows.Automation.AutomationElement]$Parent, [string]$Type)

    $cond = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
        [System.Windows.Automation.ControlType]::$Type)

    return $Parent.FindAll($scope, $cond)
}

function Invoke-Element {
    param([System.Windows.Automation.AutomationElement]$Element)

    $pattern = $Element.GetCurrentPattern(
        [System.Windows.Automation.InvokePattern]::Pattern)

    $pattern.Invoke()
}

function Click-Element {
    param([System.Windows.Automation.AutomationElement]$Element)

    $box = $Element.Current.BoundingRectangle

    if ($box.Width -le 0 -or $box.Height -le 0) {
        return
    }

    $x = [int]($box.Left + $box.Width / 2)
    $y = [int]($box.Top + $box.Height / 2)

    [void][NativeMethods]::SetCursorPos($x, $y)
    Start-Sleep -Milliseconds 250

    # Нажатие и отпускание левой кнопки мыши.
    [NativeMethods]::mouse_event(0x0002, 0, 0, 0, [UIntPtr]::Zero)
    Start-Sleep -Milliseconds 80
    [NativeMethods]::mouse_event(0x0004, 0, 0, 0, [UIntPtr]::Zero)

    Start-Sleep -Milliseconds 350
}

function Set-Text {
    param([System.Windows.Automation.AutomationElement]$Element, [string]$Text)

    $pattern = $null

    if ($Element.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern, [ref]$pattern)) {
        $pattern.SetValue($Text)
        return
    }

    # Если шаблон значения не поддерживается — ввод с клавиатуры.
    Click-Element -Element $Element

    [System.Windows.Forms.SendKeys]::SendWait("^a")
    Start-Sleep -Milliseconds 150
    [System.Windows.Forms.SendKeys]::SendWait($Text)
    Start-Sleep -Milliseconds 350
}

function Dump-Tree {
    param(
        [System.Windows.Automation.AutomationElement]$Element,
        [int]$Depth = 0,
        [int]$MaxDepth = 5
    )

    if ($null -eq $Element -or $Depth -gt $MaxDepth) { return }

    $c = $Element.Current
    $type = $c.ControlType.ProgrammaticName -replace 'ControlType\.', ''

    $line = ('  ' * $Depth) + "[$type] name='$($c.Name)' id='$($c.AutomationId)'"

    $isOn = $true
    try { $isOn = -not $c.IsOffscreen } catch { }

    if (-not $isOn) { $line += " (offscreen)" }

    Write-Output $line

    $child = $walker.GetFirstChild($Element)

    while ($null -ne $child) {
        Dump-Tree -Element $child -Depth ($Depth + 1) -MaxDepth $MaxDepth
        $child = $walker.GetNextSibling($child)
    }
}

# --- 1. Главное окно приложения ---

if ($AppProcessId -eq 0) {
    $proc = Get-Process Ktpt.CourseOrder -ErrorAction SilentlyContinue | Select-Object -First 1

    if ($null -eq $proc) {
        Write-Output "Приложение не запущено."
        exit 1
    }

    $AppProcessId = $proc.Id
}

$main = Find-Window -NameLike "Магазин автозапчастей" -OwnerProcessId $AppProcessId

if ($null -eq $main) {
    Write-Output "Главное окно не найдено (PID=$AppProcessId)."
    exit 1
}

Write-Output "Главное окно: '$($main.Current.Name)' PID=$AppProcessId"

$currentUser = Find-ById -Parent $main -AutomationId "CurrentUserText"

if ($null -ne $currentUser) {
    Write-Output "Текущий пользователь до входа: '$($currentUser.Current.Name)'"
}

if ($DumpOnly) {
    Dump-Tree -Element $main -MaxDepth 2
    exit 0
}

# --- 2. Открываем окно авторизации ---

$alreadyIn = $false

if ($null -ne $currentUser -and $currentUser.Current.Name -ne "Гость") {
    Write-Output "Пользователь уже авторизован: '$($currentUser.Current.Name)'"
    $alreadyIn = $true
}

if (-not $alreadyIn) {
    # Модальное окно авторизации входит в дерево UIA как потомок главного окна.
    $auth = Find-ChildWindow -Parent $main -NameLike "Авторизация"

    if ($null -eq $auth) {
        $accountButton = Find-ById -Parent $main -AutomationId "AccountButton"

        if ($null -eq $accountButton) {
            Write-Output "Кнопка «Вход / Регистрация» не найдена."
            exit 1
        }

        Invoke-Element -Element $accountButton

        Start-Sleep -Seconds 2

        $auth = Find-ChildWindow -Parent $main -NameLike "Авторизация"
    }

    if ($null -eq $auth) {
        Write-Output "Окно авторизации не открылось."
        exit 1
    }

    Write-Output ""
    Write-Output "--- Окно авторизации ---"
    Dump-Tree -Element $auth -MaxDepth 4

    # --- 3. Вкладка «Вход» ---

    $loginTab = Find-ByName -Parent $auth -Name "Вход" -Type "TabItem"

    if ($null -ne $loginTab) {
        try {
            $selection = $loginTab.GetCurrentPattern(
                [System.Windows.Automation.SelectionItemPattern]::Pattern)
            $selection.Select()
            Start-Sleep -Milliseconds 500
        } catch { }
    }

    # --- 4. Заполняем логин и пароль ---

    # Окно авторизации выводится на передний план: без этого
    # ввод с клавиатуры уходит в другое окно.
    Set-Foreground -Element $auth

    $edits = Find-All -Parent $auth -Type "Edit"

    Write-Output ""
    Write-Output "Текстовых полей в окне авторизации: $($edits.Count)"

    if ($edits.Count -lt 2) {
        Write-Output "Не найдены поля логина и пароля."
        exit 1
    }

    Set-Text -Element $edits[0] -Text $Login
    Start-Sleep -Milliseconds 400

    # Перед каждым следующим шагом окно переищется:
    # после ввода элемент управления может устареть.
    $auth = Find-ChildWindow -Parent $main -NameLike "Авторизация"
    $edits = Find-All -Parent $auth -Type "Edit"

    Set-Text -Element $edits[1] -Text $Password
    Start-Sleep -Milliseconds 400

    Write-Output "Логин введён: '$Login'"
    Write-Output "Пароль введён: '$Password'"

    # --- 5. Нажимаем «Войти» ---

    $auth = Find-ChildWindow -Parent $main -NameLike "Авторизация"

    if ($null -eq $auth) {
        Write-Output "Окно авторизации закрылось до нажатия «Войти»."
        exit 1
    }

    $loginButton = $null

    try {
        $loginButton = Find-ByName -Parent $auth -Name "Войти" -Type "Button"
    } catch {
        Write-Output "Кнопка «Войти» недоступна: $($_.Exception.Message)"
    }

    if ($null -eq $loginButton) {
        Write-Output "Кнопка «Войти» не найдена."
        exit 1
    }

    Invoke-Element -Element $loginButton

    Start-Sleep -Seconds 3
}

# --- 6. Проверка результата ---

Write-Output ""
Write-Output "--- Результат входа ---"

$main = Find-Window -NameLike "Магазин автозапчастей" -OwnerProcessId $AppProcessId
$currentUser = Find-ById -Parent $main -AutomationId "CurrentUserText"

if ($null -eq $currentUser) {
    Write-Output "Не удалось прочитать строку текущего пользователя."
    exit 1
}

$userName = $currentUser.Current.Name
Write-Output "Текущий пользователь: '$userName'"

foreach ($id in @("MyOrdersButton", "ManagerButton", "AdminButton", "LogoutButton")) {
    $element = Find-ById -Parent $main -AutomationId $id

    if ($null -eq $element) {
        Write-Output ("  {0,-18} — не найден" -f $id)
        continue
    }

    $visible = $true
    try { $visible = -not $element.Current.IsOffscreen } catch { }

    Write-Output ("  {0,-18} — {1}" -f $id, $(if ($visible) { "видима" } else { "скрыта" }))
}

if ($userName -eq "Гость") {
    Write-Output ""
    Write-Output "ВХОД НЕ ВЫПОЛНЕН — пользователь остался гостем."
    exit 1
}

Write-Output ""
Write-Output "ВХОД ВЫПОЛНЕН: $userName"
exit 0
