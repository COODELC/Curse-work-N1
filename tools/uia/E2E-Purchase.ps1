# Сквозная проверка покупки через интерфейс: товар → корзина → оплата → мои заказы → отмена.
# Заказ в конце отменяется через само приложение, поэтому товар возвращается на склад
# и база остаётся в исходном состоянии.
#
# Запуск:  powershell -File tools/uia/E2E-Purchase.ps1 -ProcessId 1234

param(
    [int]$ProcessId = 0
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

$root = [System.Windows.Automation.AutomationElement]::RootElement
$walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
$scope = [System.Windows.Automation.TreeScope]::Descendants

# Счётчики живут в области скрипта: обычные $passed++ внутри функции
# изменяли бы локальную копию переменной.
$script:passed = 0
$script:failed = 0

function Check {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][bool]$Condition,
        [string]$Detail = ""
    )

    if ($Condition) {
        $script:passed++
        Write-Output "  [OK]   $Name$(if ($Detail) { " — $Detail" })"
    }
    else {
        $script:failed++
        Write-Output "  [FAIL] $Name$(if ($Detail) { " — $Detail" })"
    }
}

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

function Find-TextLike {
    param([System.Windows.Automation.AutomationElement]$Parent, [string]$Like)

    $all = $Parent.FindAll($scope, [System.Windows.Automation.Condition]::TrueCondition)

    foreach ($element in $all) {
        try {
            if ($element.Current.ControlType.ProgrammaticName -notmatch 'Text') { continue }
            if ($element.Current.Name -like $Like) { return $element }
        } catch { }
    }

    return $null
}

# Список заказов виртуализирован, поэтому в дереве UIA видны только
# отображаемые строки. Настоящее количество берётся из подписи «Всего заказов: N».
function Get-OrdersCount {
    param([System.Windows.Automation.AutomationElement]$Window)

    $text = Find-TextLike -Parent $Window -Like "Всего заказов:*"

    if ($null -eq $text) { return -1 }

    $digits = ([regex]::Match($text.Current.Name, '\d+')).Value

    if ($digits -eq "") { return -1 }

    return [int]$digits
}

# Статус верхней (самой новой) строки списка.
function Get-FirstStatus {
    param([System.Windows.Automation.AutomationElement]$Window)

    $all = $Window.FindAll($scope, [System.Windows.Automation.Condition]::TrueCondition)

    foreach ($element in $all) {
        try {
            if ($element.Current.ControlType.ProgrammaticName -notmatch 'Text') { continue }
            if ($element.Current.Name -like "Статус:*") { return $element.Current.Name }
        } catch { }
    }

    return $null
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

# --- Подготовка ---

if ($ProcessId -eq 0) {
    $proc = Get-Process Ktpt.CourseOrder -ErrorAction SilentlyContinue | Select-Object -First 1

    if ($null -eq $proc) {
        Write-Output "Приложение не запущено."
        exit 1
    }

    $ProcessId = $proc.Id
}

$main = Find-Window -NameLike "Магазин автозапчастей" -OwnerProcessId $ProcessId

if ($null -eq $main) {
    Write-Output "Главное окно не найдено (PID=$ProcessId)."
    exit 1
}

Write-Output "Проверка покупки через интерфейс, PID=$ProcessId"
Write-Output ""

# --- 1. Заказов до покупки ---

Write-Output "=== 1. Количество заказов до покупки ==="

$myOrders = Find-ById -Parent $main -AutomationId "MyOrdersButton"
Invoke-Element -Element $myOrders
Start-Sleep -Seconds 2

$ordersWindow = Find-Window -NameLike "Мои заказы" -OwnerProcessId $ProcessId
$ordersBefore = Get-OrdersCount -Window $ordersWindow

Write-Output "  Заказов у пользователя: $ordersBefore"

Close-Window -Window $ordersWindow

# --- 2. Добавление товара в корзину ---

Write-Output ""
Write-Output "=== 2. Добавление товара в корзину ==="

$main = Find-Window -NameLike "Магазин автозапчастей" -OwnerProcessId $ProcessId
$addButtons = Find-AllByName -Parent $main -Name "Добавить в корзину" -Type "Button"

Check -Name "Кнопка «Добавить в корзину» есть в каталоге" -Condition ($addButtons.Count -gt 0) -Detail "кнопок: $($addButtons.Count)"

Invoke-Element -Element $addButtons[0]
Start-Sleep -Seconds 2

# --- 3. Корзина ---

Write-Output ""
Write-Output "=== 3. Корзина ==="

$cartButton = Find-ById -Parent $main -AutomationId "CartButton"
Invoke-Element -Element $cartButton
Start-Sleep -Seconds 2

$cartWindow = Find-Window -NameLike "Корзина" -OwnerProcessId $ProcessId
Check -Name "Окно корзины открылось" -Condition ($null -ne $cartWindow)

$totalText = Find-TextLike -Parent $cartWindow -Like "Итого:*"
$totalName = if ($null -ne $totalText) { $totalText.Current.Name } else { "нет текста" }

Check -Name "Итоговая сумма в корзине рассчитана" `
    -Condition ($null -ne $totalText -and $totalText.Current.Name -notlike "Итого: 0,00*") `
    -Detail $totalName

$payButton = Find-AllByName -Parent $cartWindow -Name "Оплатить" -Type "Button"
Check -Name "Кнопка «Оплатить» доступна" -Condition ($payButton.Count -eq 1) -Detail "кнопок: $($payButton.Count)"

# --- 4. Оплата ---

Write-Output ""
Write-Output "=== 4. Оплата ==="

Invoke-Element -Element $payButton[0]
Start-Sleep -Seconds 3

$statusText = Find-TextLike -Parent $cartWindow -Like "*Заказ*"
$statusName = if ($null -ne $statusText) { $statusText.Current.Name } else { "нет текста" }

Check -Name "После оплаты показано подтверждение с номером заказа" `
    -Condition ($null -ne $statusText) `
    -Detail $statusName

Close-Window -Window $cartWindow

# --- 5. Заказ появился в истории ---

Write-Output ""
Write-Output "=== 5. Заказ появился в «Моих заказах» ==="

$main = Find-Window -NameLike "Магазин автозапчастей" -OwnerProcessId $ProcessId
$myOrders = Find-ById -Parent $main -AutomationId "MyOrdersButton"
Invoke-Element -Element $myOrders
Start-Sleep -Seconds 2

$ordersWindow = Find-Window -NameLike "Мои заказы" -OwnerProcessId $ProcessId
$ordersAfter = Get-OrdersCount -Window $ordersWindow

Check -Name "Количество заказов увеличилось на 1" `
    -Condition ($ordersAfter -eq ($ordersBefore + 1)) `
    -Detail "было $ordersBefore, стало $ordersAfter"

$newStatusName = Get-FirstStatus -Window $ordersWindow

Check -Name "Новый заказ (верхний в списке) в статусе «Оформление заказа»" `
    -Condition ($newStatusName -eq "Статус: Оформление заказа") `
    -Detail $(if ($null -ne $newStatusName) { $newStatusName } else { "нет текста" })

# --- 6. Отмена заказа возвращает товар ---

Write-Output ""
Write-Output "=== 6. Отмена заказа ==="

$cancelButtons = Find-AllByName -Parent $ordersWindow -Name "Отменить" -Type "Button"
Invoke-Element -Element $cancelButtons[0]
Start-Sleep -Seconds 3

$ordersWindow = Find-Window -NameLike "Мои заказы" -OwnerProcessId $ProcessId
$cancelledName = Get-FirstStatus -Window $ordersWindow

Check -Name "Заказ переведён в статус «Отменён»" `
    -Condition ($cancelledName -eq "Статус: Отменён") `
    -Detail $(if ($null -ne $cancelledName) { $cancelledName } else { "нет текста" })

Close-Window -Window $ordersWindow

Write-Output ""
Write-Output "ИТОГО: успешно $script:passed, провалено $script:failed"

if ($script:failed -eq 0) { exit 0 } else { exit 1 }
