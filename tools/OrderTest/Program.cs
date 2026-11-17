// Сквозная проверка работы с заказами на живой базе cursevork.
// Тест работает через настоящий OrderService приложения и после прогона
// полностью откатывает все изменения, возвращая базу в исходное состояние.
//
// Запуск:  dotnet run tools/OrderTest/OrderTest.cs

using System.Text;
using Ktpt.CourseOrder.Auth;
using Ktpt.CourseOrder.Data;
using Ktpt.CourseOrder.Services;
using Microsoft.EntityFrameworkCore;

Console.OutputEncoding = Encoding.UTF8;

var repoRoot = FindRepoRoot();
var connectionFile = Path.Combine(repoRoot, "scr", "Ktpt.CourseOrder", "Connection.txt");
var connectionString = File.ReadAllText(connectionFile).Trim();

var factory = new DbFactory(connectionString);
var service = new OrderService(factory);

var passed = 0;
var failed = 0;

void Check(string name, bool condition, string detail = "")
{
    if (condition)
    {
        passed++;
        Console.WriteLine($"  [OK]   {name}{(detail.Length > 0 ? " — " + detail : "")}");
    }
    else
    {
        failed++;
        Console.WriteLine($"  [FAIL] {name}{(detail.Length > 0 ? " — " + detail : "")}");
    }
}

void Section(string title) => Console.WriteLine($"\n=== {title} ===");

// ---------- Снимок исходного состояния ----------

int userId;
var stockBefore = new Dictionary<int, int>();
var ordersBefore = new List<int>();
var statuses = new Dictionary<string, int>();

using (var db = factory.Create())
{
    var maria = await db.Users.FirstOrDefaultAsync(u => u.Login == "maria");

    if (maria is null)
    {
        Console.Error.WriteLine("Пользователь maria не найден — нечего тестировать.");
        Environment.Exit(1);
    }

    userId = maria.Id_User;

    foreach (var item in await db.Items.ToListAsync())
    {
        stockBefore[item.Id_Item] = item.Quantity;
    }

    ordersBefore = await db.Orders.Select(o => o.Id_Order).ToListAsync();

    foreach (var status in await db.OrderStatuses.ToListAsync())
    {
        statuses[status.Name_Status] = status.Id_status;
    }
}

Console.WriteLine($"Подключено: {connectionString.Split(';')[0]}");
Console.WriteLine($"Пользователь для тестов: maria (Id={userId}), товаров: {stockBefore.Count}, заказов: {ordersBefore.Count}");

var createdOrderIds = new List<int>();

try
{
    // ---------- 0. Авторизация администратора ----------

    Section("0. Авторизация администратора");

    var authService = new AuthService(factory);

    var wrongPassword = await authService.LoginAsync("ander", "неверный-пароль");
    Check("Вход с неверным паролем отклоняется", wrongPassword is null);

    var emptyPassword = await authService.LoginAsync("ander", "");
    Check("Вход с пустым паролем отклоняется", emptyPassword is null);

    var unknownUser = await authService.LoginAsync("нет-такого", "8942");
    Check("Вход с неизвестным логином отклоняется", unknownUser is null);

    var admin = await authService.LoginAsync("ander", "8942");
    Check("Вход администратора (ander) выполняется", admin is not null);

    if (admin is not null)
    {
        Check("Роль администратора определена", admin.Role == UserRole.Administrator, $"роль: {admin.Role}");
        Check("Имя администратора", admin.FirstName == "Андрей", $"{admin.FirstName} {admin.SecondName}");
        Check("Право управления товарами", Permissions.Allows(admin.Role, Permission.ManageProducts));
        Check("Право управления заказами", Permissions.Allows(admin.Role, Permission.ManageOrders));
        Check("Право управления пользователями", Permissions.Allows(admin.Role, Permission.ManageUsers));
        Check("Право оставлять отзывы", Permissions.Allows(admin.Role, Permission.AddReview));
    }

    var customer = await authService.LoginAsync("maria", "1234");
    Check("Вход покупателя (maria) выполняется", customer is not null);

    if (customer is not null)
    {
        Check("Роль покупателя определена", customer.Role == UserRole.Customer, $"роль: {customer.Role}");
        Check("У покупателя нет прав управления товарами",
            !Permissions.Allows(customer.Role, Permission.ManageProducts));
        Check("У покупателя нет прав управления пользователями",
            !Permissions.Allows(customer.Role, Permission.ManageUsers));
    }

    // ---------- 1. Проверки входных данных ----------

    Section("1. Оформление заказа: проверка входных данных");

    var empty = await service.CreateOrderAsync(userId, new List<(int, int)>());
    Check("Пустая корзина отклоняется", !empty.Success && empty.Message.Contains("пуста"), empty.Message);

    var missing = await service.CreateOrderAsync(userId, new List<(int, int)> { (999999, 1) });
    Check("Несуществующий товар отклоняется", !missing.Success && missing.Message.Contains("не найден"), missing.Message);

    var zeroQty = await service.CreateOrderAsync(userId, new List<(int, int)> { (8, 0) });
    Check("Количество 0 отклоняется", !zeroQty.Success && zeroQty.Message.Contains("больше нуля"), zeroQty.Message);

    var negativeQty = await service.CreateOrderAsync(userId, new List<(int, int)> { (8, -3) });
    Check("Отрицательное количество отклоняется", !negativeQty.Success, negativeQty.Message);

    // Товар с заведомо малым остатком — заказываем больше, чем есть.
    int scarceItemId = 0;
    foreach (var pair in stockBefore)
    {
        if (pair.Value > 0 && (scarceItemId == 0 || pair.Value < stockBefore[scarceItemId]))
        {
            scarceItemId = pair.Key;
        }
    }

    var tooMany = await service.CreateOrderAsync(userId, new List<(int, int)> { (scarceItemId, stockBefore[scarceItemId] + 50) });
    Check("Заказ сверх остатка отклоняется", !tooMany.Success && tooMany.Message.Contains("Недостаточно"), tooMany.Message);

    int stockAfterFailed;
    using (var db = factory.Create())
    {
        stockAfterFailed = await db.Items.Where(i => i.Id_Item == scarceItemId).Select(i => i.Quantity).FirstAsync();
    }

    Check("Неудачный заказ не списывает товар", stockAfterFailed == stockBefore[scarceItemId],
        $"остаток {stockAfterFailed}, ожидалось {stockBefore[scarceItemId]}");

    // ---------- 2. Скидка постоянного покупателя ----------

    Section("2. Скидка постоянного покупателя (25%)");

    // Item 8 (Свеча зажигания): maria не заказывала его в прошлом месяце → скидка.
    // Item 2 (Воздушный фильтр): maria заказывала его в октябре → без скидки.
    int discountItemId = 8;
    int noDiscountItemId = 2;

    decimal priceDiscountItem, priceNoDiscountItem;
    using (var db = factory.Create())
    {
        priceDiscountItem = await db.Items.Where(i => i.Id_Item == discountItemId).Select(i => i.Price).FirstAsync();
        priceNoDiscountItem = await db.Items.Where(i => i.Id_Item == noDiscountItemId).Select(i => i.Price).FirstAsync();
    }

    var withDiscount = await service.CreateOrderAsync(userId, new List<(int, int)> { (discountItemId, 1) });
    Check("Заказ без прошлой покупки принимается", withDiscount.Success, withDiscount.Message);

    if (withDiscount.Success)
    {
        createdOrderIds.Add(withDiscount.OrderId);

        var expected = priceDiscountItem * 0.75m;
        Check("Применена скидка 25% к новому товару",
            Math.Abs(withDiscount.Total - expected) < 0.01m,
            $"итого {withDiscount.Total:N2}, ожидалось {expected:N2}");
        Check("Сумма скидки посчитана", Math.Abs(withDiscount.DiscountTotal - (priceDiscountItem - expected)) < 0.01m,
            $"скидка {withDiscount.DiscountTotal:N2}");
    }

    var withoutDiscount = await service.CreateOrderAsync(userId, new List<(int, int)> { (noDiscountItemId, 1) });
    Check("Повторный заказ товара прошлого месяца принимается", withoutDiscount.Success, withoutDiscount.Message);

    if (withoutDiscount.Success)
    {
        createdOrderIds.Add(withoutDiscount.OrderId);

        Check("Скидка не применяется к товару, заказанному в прошлом месяце",
            Math.Abs(withoutDiscount.Total - priceNoDiscountItem) < 0.01m,
            $"итого {withoutDiscount.Total:N2}, ожидалось {priceNoDiscountItem:N2}");
    }

    // ---------- 3. Списание остатка и строки заказа ----------

    Section("3. Списание остатка и строки заказа");

    int stockNow;
    int rowsNow;
    using (var db = factory.Create())
    {
        stockNow = await db.Items.Where(i => i.Id_Item == discountItemId).Select(i => i.Quantity).FirstAsync();

        rowsNow = await db.Orders
            .Where(o => createdOrderIds.Contains(o.Id_Order))
            .CountAsync();
    }

    Check("Остаток уменьшился на 1", stockNow == stockBefore[discountItemId] - 1,
        $"остаток {stockNow}, было {stockBefore[discountItemId]}");

    Check("Создана строка заказа", rowsNow == createdOrderIds.Count, $"строк: {rowsNow}");

    // ---------- 4. Количество разворачивается в отдельные строки ----------

    Section("4. Несколько единиц = несколько строк заказа");

    int multiItemId = 8;
    var multi = await service.CreateOrderAsync(userId, new List<(int, int)> { (multiItemId, 3) });
    Check("Заказ нескольких единиц принимается", multi.Success, multi.Message);

    if (multi.Success)
    {
        createdOrderIds.Add(multi.OrderId);

        Check("Позиций в заказе = 3", multi.PositionsCount == 3, $"позиций: {multi.PositionsCount}");
        Check("Заказ признан крупным (>5)", !multi.IsLarge, "3 позиции — не крупный");

        int stockAfterMulti;
        using (var db = factory.Create())
        {
            stockAfterMulti = await db.Items.Where(i => i.Id_Item == multiItemId).Select(i => i.Quantity).FirstAsync();
        }

        Check("Списано 3 единицы", stockAfterMulti == stockBefore[multiItemId] - 4,
            $"остаток {stockAfterMulti}, ожидалось {stockBefore[multiItemId] - 4}");

        var big = await service.CreateOrderAsync(userId, new List<(int, int)> { (8, 6) });
        Check("Крупный заказ (6 единиц) принимается", big.Success, big.Message);

        if (big.Success)
        {
            createdOrderIds.Add(big.OrderId);

            Check("Крупный заказ помечен флагом IsLarge", big.IsLarge, $"позиций: {big.PositionsCount}");
            Check("Сумма крупного заказа посчитана", big.Total > 0, $"{big.Total:N2} руб.");
        }
    }

    // ---------- 5. Смена статуса ----------

    Section("5. Изменение статуса заказа");

    int testOrderId = createdOrderIds.First();

    using (var db = factory.Create())
    {
        var current = await db.Orders.Where(o => o.Id_Order == testOrderId).Select(o => o.Status_Id).FirstAsync();
        Check("Новый заказ в статусе «Оформление заказа»",
            current == statuses["Оформление заказа"], $"статус Id={current}");
    }

    await service.SetStatusAsync(testOrderId, statuses["Собирается"]);

    using (var db = factory.Create())
    {
        var changed = await db.Orders.Where(o => o.Id_Order == testOrderId).Select(o => o.Status_Id).FirstAsync();
        Check("Статус изменён на «Собирается»", changed == statuses["Собирается"], $"статус Id={changed}");
    }

    try
    {
        await service.SetStatusAsync(999999, statuses["Собирается"]);
        Check("Смена статуса несуществующего заказа падает с исключением", false);
    }
    catch (InvalidOperationException ex)
    {
        Check("Смена статуса несуществующего заказа падает с исключением", ex.Message.Contains("не найден"), ex.Message);
    }

    // ---------- 6. Отмена заказа и возврат товара ----------

    Section("6. Отмена заказа и возврат товара на склад");

    int stockBeforeCancel;
    using (var db = factory.Create())
    {
        stockBeforeCancel = await db.Items.Where(i => i.Id_Item == discountItemId).Select(i => i.Quantity).FirstAsync();
    }

    var cancelled = await service.CancelOrderAsync(testOrderId);
    Check("Отмена заказа возвращает true", cancelled);

    int stockAfterCancel;
    int statusAfterCancel;
    using (var db = factory.Create())
    {
        stockAfterCancel = await db.Items.Where(i => i.Id_Item == discountItemId).Select(i => i.Quantity).FirstAsync();
        statusAfterCancel = await db.Orders.Where(o => o.Id_Order == testOrderId).Select(o => o.Status_Id).FirstAsync();
    }

    Check("Товар возвращён на склад (+1)", stockAfterCancel == stockBeforeCancel + 1,
        $"остаток {stockAfterCancel}, было {stockBeforeCancel}");
    Check("Статус заказа — «Отменён»", statusAfterCancel == statuses["Отменён"], $"статус Id={statusAfterCancel}");

    var cancelledAgain = await service.CancelOrderAsync(testOrderId);
    Check("Повторная отмена возвращает false", !cancelledAgain);

    int stockAfterSecondCancel;
    using (var db = factory.Create())
    {
        stockAfterSecondCancel = await db.Items.Where(i => i.Id_Item == discountItemId).Select(i => i.Quantity).FirstAsync();
    }

    Check("Повторная отмена НЕ возвращает товар второй раз", stockAfterSecondCancel == stockAfterCancel,
        $"остаток {stockAfterSecondCancel}, ожидалось {stockAfterCancel}");

    try
    {
        await service.CancelOrderAsync(999999);
        Check("Отмена несуществующего заказа падает с исключением", false);
    }
    catch (InvalidOperationException ex)
    {
        Check("Отмена несуществующего заказа падает с исключением", ex.Message.Contains("не найден"), ex.Message);
    }

    // ---------- 7. Просмотр истории заказов ----------

    Section("7. Просмотр истории заказов");

    var userOrders = await service.GetUserOrdersAsync(userId);
    Check("История заказов пользователя загружается", userOrders.Count > 0, $"заказов: {userOrders.Count}");
    Check("История отсортирована по дате (новые сверху)",
        userOrders.Zip(userOrders.Skip(1)).All(p => p.First.Date >= p.Second.Date));
    Check("В истории есть созданные тестом заказы",
        createdOrderIds.Any(id => userOrders.Any(o => o.Id == id)));

    var allOrders = await service.GetAllOrdersAsync();
    Check("Список всех заказов загружается", allOrders.Count >= ordersBefore.Count,
        $"заказов: {allOrders.Count}");
    Check("В общем списке заполнено имя покупателя", allOrders.All(o => o.UserName.Length > 0));

    var statusOptions = await service.LoadStatusesAsync();
    Check("Список статусов для выбора загружается", statusOptions.Count == statuses.Count, $"статусов: {statusOptions.Count}");
    Check("В списке статусов есть «Отменён»", statusOptions.Any(s => s.Name == "Отменён"));
}
finally
{
    // ---------- Откат: возвращаем базу в исходное состояние ----------

    Section("Откат изменений");

    using (var db = factory.Create())
    {
        // Удаляются все строки заказов, созданные тестом, а не только первые.
        var testOrders = await db.Orders
            .Where(o => !ordersBefore.Contains(o.Id_Order))
            .ToListAsync();

        if (testOrders.Count > 0)
        {
            db.Orders.RemoveRange(testOrders);
            await db.SaveChangesAsync();

            Console.WriteLine($"  Удалено тестовых заказов: {testOrders.Count}");
        }
    }

    using (var db = factory.Create())
    {
        foreach (var pair in stockBefore)
        {
            var item = await db.Items.FirstOrDefaultAsync(i => i.Id_Item == pair.Key);

            if (item is not null && item.Quantity != pair.Value)
            {
                item.Quantity = pair.Value;
                Console.WriteLine($"  Восстановлен остаток {pair.Key}: {pair.Value}");
            }
        }

        await db.SaveChangesAsync();
    }

    using (var db = factory.Create())
    {
        var ordersNow = await db.Orders.Select(o => o.Id_Order).OrderBy(i => i).ToListAsync();
        var sameOrders = ordersNow.SequenceEqual(ordersBefore.OrderBy(i => i));

        var stockRestored = true;

        foreach (var pair in stockBefore)
        {
            var quantity = await db.Items.Where(i => i.Id_Item == pair.Key).Select(i => i.Quantity).FirstAsync();

            if (quantity != pair.Value)
            {
                stockRestored = false;
            }
        }

        Check("Состав заказов восстановлен", sameOrders);
        Check("Остатки товаров восстановлены", stockRestored);
    }
}

Console.WriteLine();
Console.WriteLine($"ИТОГО: успешно {passed}, провалено {failed}");

Environment.Exit(failed == 0 ? 0 : 1);

static string FindRepoRoot()
{
    var dir = new DirectoryInfo(Directory.GetCurrentDirectory());

    while (dir is not null)
    {
        if (Directory.Exists(Path.Combine(dir.FullName, "sql")) &&
            Directory.Exists(Path.Combine(dir.FullName, "scr")))
        {
            return dir.FullName;
        }

        dir = dir.Parent;
    }

    return Directory.GetCurrentDirectory();
}
