# Снимает экранные формы приложения «Магазин автозапчастей».
#   powershell -File tools/uia/Capture-Forms.ps1 -ProcessId 24180 -OutputDir docs\screens

param(
    [Parameter(Mandatory = $true)][int]$ProcessId,
    [string]$OutputDir = "docs\screens"
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

$root = [System.Windows.Automation.AutomationElement]::RootElement
$walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
$scope = [System.Windows.Automation.TreeScope]::Descendants

function Get-MainWindow {
    $child = $walker.GetFirstChild($root)

    while ($null -ne $child) {
        try {
            if ($child.Current.ProcessId -eq $ProcessId -and $child.Current.Name -like '*Магазин автозапчастей*') {
                return $child
            }
        } catch { }

        $child = $walker.GetNextSibling($child)
    }

    return $null
}

function Find-ById {
    param([System.Windows.Automation.AutomationElement]$Parent, [string]$AutomationId)

    $cond = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::AutomationIdProperty, $AutomationId)

    return $Parent.FindFirst($scope, $cond)
}

function Find-ByName {
    param(
        [System.Windows.Automation.AutomationElement]$Parent,
        [string]$NameLike,
        [string]$Type = ""
    )

    # PropertyCondition ищет точное совпадение, поэтому имя фильтруется вручную.
    $all = $Parent.FindAll($scope, [System.Windows.Automation.Condition]::TrueCondition)

    foreach ($element in $all) {
        try {
            if ($element.Current.Name -notlike $NameLike) { continue }

            if ($Type -ne "") {
                $controlType = $element.Current.ControlType.ProgrammaticName -replace 'ControlType\.', ''

                if ($controlType -ne $Type) { continue }
            }

            return $element
        } catch { }
    }

    return $null
}

function Invoke-Element {
    param([System.Windows.Automation.AutomationElement]$Element)

    $pattern = $Element.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
    $pattern.Invoke()
}

function Save-Shot {
    param([string]$TitleLike, [string]$FileName)

    $script = Join-Path $PSScriptRoot 'Screenshot.ps1'

    & powershell -NoProfile -ExecutionPolicy Bypass -File $script `
        -ProcessId $ProcessId -TitleLike $TitleLike `
        -Output (Join-Path $OutputDir $FileName) -Width 1200 -Height 740
}

function Close-Window {
    param([System.Windows.Automation.AutomationElement]$Window)

    try {
        $pattern = $Window.GetCurrentPattern(
            [System.Windows.Automation.WindowPattern]::Pattern)

        $pattern.Close()
    } catch {
        # Если закрыть не удалось, окно закрывается кнопкой.
        $button = Find-ByName -Parent $Window -NameLike 'Закрыть*' -Type 'Button'

        if ($null -eq $button) {
            $button = Find-ByName -Parent $Window -NameLike 'Отмена*' -Type 'Button'
        }

        if ($null -ne $button) { Invoke-Element -Element $button }
    }

    Start-Sleep -Milliseconds 800
}

function Close-Modal {
    param([string]$TitleLike)

    # Модальное окно закрывается, чтобы можно было открыть следующее.
    $child = $walker.GetFirstChild($root)

    while ($null -ne $child) {
        try {
            if ($child.Current.ProcessId -eq $ProcessId -and $child.Current.Name -like "*$TitleLike*") {
                Close-Window -Window $child

                return
            }
        } catch { }

        $child = $walker.GetNextSibling($child)
    }
}

function Close-StaleModals {
    # Окна, оставшиеся от прошлого прогона, блокируют главное окно.
    # Модальные диалоги входят в дерево как потомки главного окна.
    $winCond = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
        [System.Windows.Automation.ControlType]::Window)

    $main = Get-MainWindow

    if ($null -ne $main) {
        foreach ($window in $main.FindAll($scope, $winCond)) {
            try {
                if ($window.Current.Name -like '*Магазин автозапчастей*') { continue }

                Close-Window -Window $window
            } catch { }
        }
    }

    $child = $walker.GetFirstChild($root)

    while ($null -ne $child) {
        try {
            $name = $child.Current.Name

            if ($child.Current.ProcessId -eq $ProcessId -and $name -notlike '*Магазин автозапчастей*') {
                Close-Window -Window $child
            }
        } catch { }

        $child = $walker.GetNextSibling($child)
    }
}

if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
}

$main = Get-MainWindow

if ($null -eq $main) {
    Write-Output "Главное окно не найдено (PID=$ProcessId)."
    exit 1
}

$currentUser = Find-ById -Parent $main -AutomationId "CurrentUserText"
Write-Output "Пользователь: '$($currentUser.Current.Name)'"

# Закрываются окна, оставшиеся от предыдущего прогона.
Close-StaleModals
Start-Sleep -Milliseconds 500

# 1. Главное окно
Save-Shot -TitleLike 'Магазин автозапчастей' -FileName '01_main.png'

# 2. Описание товара — модальное окно
$details = Find-ByName -Parent $main -NameLike 'Описание товара' -Type 'Button'

if ($null -ne $details) {
    Invoke-Element -Element $details
    Start-Sleep -Seconds 2
    Save-Shot -TitleLike 'О товаре' -FileName '02_product_details.png'
    Close-Modal -TitleLike 'О товаре'
} else {
    Write-Output "Кнопка «Описание товара» не найдена."
}

# 3. Отзывы о товаре с авторами — модальное окно
$reviews = Find-ByName -Parent $main -NameLike 'Отзывов*' -Type 'Button'

if ($null -ne $reviews) {
    Invoke-Element -Element $reviews
    Start-Sleep -Seconds 2
    Save-Shot -TitleLike 'Отзывы о товаре' -FileName '03_reviews.png'
    Close-Modal -TitleLike 'Отзывы о товаре'
} else {
    Write-Output "Ссылка на отзывы не найдена."
}

# 4. Обработка заказов
$manager = Find-ById -Parent $main -AutomationId "ManagerButton"

if ($null -ne $manager) {
    Invoke-Element -Element $manager
    Start-Sleep -Seconds 2
    Save-Shot -TitleLike 'Обработка заказов' -FileName '04_manager_orders.png'
} else {
    Write-Output "Кнопка «Заказы» недоступна."
}

# 5. Управление магазином
$admin = Find-ById -Parent $main -AutomationId "AdminButton"

if ($null -ne $admin) {
    Invoke-Element -Element $admin
    Start-Sleep -Seconds 2
    Save-Shot -TitleLike 'Управление магазином' -FileName '05_admin.png'
}

# 6. Мои заказы
$myOrders = Find-ById -Parent $main -AutomationId "MyOrdersButton"

if ($null -ne $myOrders) {
    Invoke-Element -Element $myOrders
    Start-Sleep -Seconds 2
    Save-Shot -TitleLike 'Мои заказы' -FileName '06_my_orders.png'
}

# 7. Корзина
$cart = Find-ById -Parent $main -AutomationId "CartButton"

if ($null -ne $cart) {
    Invoke-Element -Element $cart
    Start-Sleep -Seconds 2
    Save-Shot -TitleLike 'Корзина' -FileName '07_cart.png'
}

# 8. Профиль — модальное окно
$profile = Find-ById -Parent $main -AutomationId "ProfileButton"

if ($null -ne $profile) {
    Invoke-Element -Element $profile
    Start-Sleep -Seconds 2
    Save-Shot -TitleLike 'Профиль' -FileName '08_profile.png'
    Close-Modal -TitleLike 'Профиль'
}

# 9. Роли — модальное окно
$roles = Find-ById -Parent $main -AutomationId "RolesButton"

if ($null -ne $roles) {
    Invoke-Element -Element $roles
    Start-Sleep -Seconds 2
    Save-Shot -TitleLike 'Управление ролями' -FileName '09_roles.png'
    Close-Modal -TitleLike 'Управление ролями'
}

Write-Output "Готово. Снимки в $OutputDir"
