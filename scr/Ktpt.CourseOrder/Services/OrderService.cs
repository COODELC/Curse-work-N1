using Ktpt.CourseOrder.Data;
using Ktpt.CourseOrder.Data.Entities;
using Microsoft.EntityFrameworkCore;

namespace Ktpt.CourseOrder.Services;

/// <summary>
/// Работа с заказами: оформление, скидка постоянному покупателю,
/// изменение статуса, отмена заказа с возвратом товара на склад.
/// </summary>
public sealed class OrderService
{
    /// <summary>Размер скидки постоянного покупателя, процентов.</summary>
    public const decimal DiscountPercent = 25m;

    /// <summary>Заказ считается «много», если количество позиций больше этого значения.</summary>
    public const int LargeOrderThreshold = 5;

    private readonly DbFactory _factory;

    public OrderService(DbFactory factory)
    {
        _factory = factory;
    }

    /// <summary>
    /// Оформляет заказ: проверяет остатки, применяет скидку,
    /// списывает товар и сохраняет строки заказа в одной транзакции.
    /// </summary>
    public async Task<OrderResult> CreateOrderAsync(
        int userId,
        IReadOnlyList<(int ItemId, int Quantity)> lines)
    {
        if (lines.Count == 0)
        {
            return new OrderResult
            {
                Success = false,
                Message = "Корзина пуста."
            };
        }

        using var db = _factory.Create();

        var now = DateTime.Now;

        var previousMonthStart = new DateTime(now.Year, now.Month, 1).AddMonths(-1);
        var previousMonthEnd = new DateTime(now.Year, now.Month, 1);

        // Товары, которые покупатель уже заказывал в прошлом календарном месяце.
        var purchasedItemIds = await db.Orders
            .AsNoTracking()
            .Where(order =>
                order.UserId == userId
                && order.Date_Order >= previousMonthStart
                && order.Date_Order < previousMonthEnd)
            .Select(order => order.Item_Id)
            .Distinct()
            .ToListAsync();

        var itemIds = lines.Select(line => line.ItemId).Distinct().ToList();

        var items = await db.Items
            .Where(item => itemIds.Contains(item.Id_Item))
            .ToDictionaryAsync(item => item.Id_Item);

        // Первая проверка: заказ нельзя оформить, если позиции нет в каталоге.
        foreach (var (itemId, quantity) in lines)
        {
            if (!items.TryGetValue(itemId, out var item))
            {
                return new OrderResult
                {
                    Success = false,
                    Message = $"Товар с идентификатором {itemId} не найден в каталоге."
                };
            }

            if (quantity <= 0)
            {
                return new OrderResult
                {
                    Success = false,
                    Message = "Количество товара должно быть больше нуля."
                };
            }

            if (item.Quantity < quantity)
            {
                return new OrderResult
                {
                    Success = false,
                    Message =
                        $"Недостаточно товара «{item.Name_Item}». " +
                        $"В наличии: {item.Quantity} шт."
                };
            }
        }

        var status = await db.OrderStatuses
            .FirstOrDefaultAsync(item => item.Name_Status == "Оформление заказа");

        if (status is null)
        {
            return new OrderResult
            {
                Success = false,
                Message = "В таблице OrderStatus нет статуса «Оформление заказа»."
            };
        }

        var total = 0m;
        var discountTotal = 0m;

        var createdOrders = new List<Orders>();

        await using var transaction = await db.Database.BeginTransactionAsync();

        foreach (var (itemId, quantity) in lines)
        {
            var item = items[itemId];

            // Скидка применяется, если в прошлом месяце по этому товару заказов не было.
            var discount = purchasedItemIds.Contains(itemId) ? 0m : DiscountPercent;

            var price = item.Price * (1 - discount / 100m);

            item.Quantity -= quantity;

            // В таблице Orders каждая запись — одна позиция заказа,
            // поэтому количество товара разворачивается в отдельные записи.
            for (var unit = 0; unit < quantity; unit++)
            {
                var order = new Orders
                {
                    Date_Order = now,
                    Item_Id = itemId,
                    UserId = userId,
                    Status_Id = status.Id_status
                };

                db.Orders.Add(order);

                createdOrders.Add(order);
            }

            total += price * quantity;
            discountTotal += (item.Price - price) * quantity;
        }

        await db.SaveChangesAsync();

        await transaction.CommitAsync();

        var orderId = createdOrders.Min(order => order.Id_Order);

        return new OrderResult
        {
            Success = true,
            OrderId = orderId,
            Total = total,
            DiscountTotal = discountTotal,
            IsLarge = lines.Sum(line => line.Quantity) > LargeOrderThreshold,
            PositionsCount = createdOrders.Count,
            Message = $"Заказ № {orderId} оформлен. Позиций: {createdOrders.Count}. Сумма: {total:N2} руб."
        };
    }

    /// <summary>
    /// История заказов пользователя.
    /// </summary>
    public async Task<List<OrderInfo>> GetUserOrdersAsync(int userId)
    {
        using var db = _factory.Create();

        var orders = await db.Orders
            .AsNoTracking()
            .Include(order => order.Item)
            .Include(order => order.Status)
            .Where(order => order.UserId == userId)
            .OrderByDescending(order => order.Date_Order)
            .ThenByDescending(order => order.Id_Order)
            .ToListAsync();

        return orders.Select(Map).ToList();
    }

    /// <summary>
    /// Все заказы — для сотрудника и администратора.
    /// </summary>
    public async Task<List<OrderInfo>> GetAllOrdersAsync()
    {
        using var db = _factory.Create();

        var orders = await db.Orders
            .AsNoTracking()
            .Include(order => order.Item)
            .Include(order => order.Status)
            .Include(order => order.User)
            .OrderByDescending(order => order.Date_Order)
            .ThenByDescending(order => order.Id_Order)
            .ToListAsync();

        return orders.Select(Map).ToList();
    }

    /// <summary>
    /// Изменение статуса заказа.
    /// </summary>
    public async Task SetStatusAsync(int orderId, int statusId)
    {
        using var db = _factory.Create();

        var order = await db.Orders
            .FirstOrDefaultAsync(item => item.Id_Order == orderId);

        if (order is null)
        {
            throw new InvalidOperationException($"Заказ № {orderId} не найден.");
        }

        order.Status_Id = statusId;

        await db.SaveChangesAsync();
    }

    /// <summary>
    /// Отмена заказа: остаток возвращается на склад.
    /// Возвращает false, если заказ уже был отменён.
    /// </summary>
    public async Task<bool> CancelOrderAsync(int orderId)
    {
        using var db = _factory.Create();

        var cancelled = await db.OrderStatuses
            .FirstOrDefaultAsync(item => item.Name_Status == "Отменён");

        if (cancelled is null)
        {
            throw new InvalidOperationException(
                "В таблице OrderStatus нет статуса «Отменён».");
        }

        var order = await db.Orders
            .Include(item => item.Item)
            .FirstOrDefaultAsync(item => item.Id_Order == orderId);

        if (order is null)
        {
            throw new InvalidOperationException($"Заказ № {orderId} не найден.");
        }

        if (order.Status_Id == cancelled.Id_status)
        {
            return false;
        }

        order.Item.Quantity += 1;

        order.Status_Id = cancelled.Id_status;

        await db.SaveChangesAsync();

        return true;
    }

    /// <summary>
    /// Список статусов заказа для выбора в интерфейсе.
    /// </summary>
    public async Task<List<OrderStatusOption>> LoadStatusesAsync()
    {
        using var db = _factory.Create();

        var statuses = await db.OrderStatuses
            .AsNoTracking()
            .OrderBy(status => status.Id_status)
            .ToListAsync();

        return statuses
            .Select(status => new OrderStatusOption
            {
                Id = status.Id_status,
                Name = status.Name_Status
            })
            .ToList();
    }

    private static OrderInfo Map(Orders order)
    {
        return new OrderInfo
        {
            Id = order.Id_Order,
            Date = order.Date_Order,
            ItemName = order.Item?.Name_Item ?? "",
            Articule = order.Item?.NameArticule ?? "",
            Price = order.Item?.Price ?? 0m,
            Quantity = 1,
            StatusId = order.Status_Id,
            Status = order.Status?.Name_Status ?? "",
            UserId = order.UserId,
            UserName = order.User is null
                ? ""
                : $"{order.User.FirstName} {order.User.SecondName}".Trim()
        };
    }
}

/// <summary>
/// Строка заказа для интерфейса.
/// </summary>
public sealed class OrderInfo
{
    public int Id { get; set; }

    public DateTime Date { get; set; }

    public string ItemName { get; set; } = "";

    public string Articule { get; set; } = "";

    public decimal Price { get; set; }

    public int Quantity { get; set; }

    public int StatusId { get; set; }

    public string Status { get; set; } = "";

    public int UserId { get; set; }

    public string UserName { get; set; } = "";

    public decimal Total => Price * Quantity;
}

/// <summary>
/// Статус заказа в виде, пригодном для выбора в интерфейсе.
/// </summary>
public sealed class OrderStatusOption
{
    public int Id { get; set; }

    public string Name { get; set; } = "";

    public override string ToString() => Name;
}

/// <summary>
/// Результат оформления заказа.
/// </summary>
public sealed class OrderResult
{
    public bool Success { get; set; }

    public int OrderId { get; set; }

    public decimal Total { get; set; }

    public decimal DiscountTotal { get; set; }

    /// <summary>Количество позиций (единиц товара) в заказе.</summary>
    public int PositionsCount { get; set; }

    /// <summary>true, если заказ «много» (позиций больше пяти).</summary>
    public bool IsLarge { get; set; }

    public string Message { get; set; } = "";
}
