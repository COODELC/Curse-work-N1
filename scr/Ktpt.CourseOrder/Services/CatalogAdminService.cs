using Ktpt.CourseOrder.Data;
using Ktpt.CourseOrder.Data.Entities;
using Microsoft.EntityFrameworkCore;

namespace Ktpt.CourseOrder.Services;

/// <summary>
/// Управление каталогом и пользователями — операции администратора.
/// </summary>
public sealed class CatalogAdminService
{
    private readonly DbFactory _factory;

    public CatalogAdminService(DbFactory factory)
    {
        _factory = factory;
    }

    /// <summary>
    /// Добавляет товар в каталог.
    /// </summary>
    public async Task AddItemAsync(
        string name,
        string articule,
        int categoryId,
        decimal price,
        int quantity)
    {
        name = name.Trim();

        if (name.Length is < 2 or > 80)
        {
            throw new ArgumentException("Наименование товара должно содержать от 2 до 80 символов.");
        }

        if (price < 0)
        {
            throw new ArgumentException("Цена товара не может быть отрицательной.");
        }

        if (quantity < 0)
        {
            throw new ArgumentException("Остаток товара не может быть отрицательным.");
        }

        using var db = _factory.Create();

        var categoryExists = await db.Categories
            .AnyAsync(category => category.Id_Category == categoryId);

        if (!categoryExists)
        {
            throw new ArgumentException("Выбрана категория, которой нет в справочнике.");
        }

        db.Items.Add(new Item
        {
            Name_Item = name,
            NameArticule = string.IsNullOrWhiteSpace(articule) ? null : articule.Trim(),
            Category_Id = categoryId,
            Price = price,
            Quantity = quantity,
            ImagePath = "/Assets/Products/no-image.png"
        });

        await db.SaveChangesAsync();
    }

    /// <summary>
    /// Изменяет параметры товара.
    /// </summary>
    public async Task UpdateItemAsync(
        int itemId,
        string name,
        string articule,
        int categoryId,
        decimal price,
        int quantity)
    {
        name = name.Trim();

        if (name.Length is < 2 or > 80)
        {
            throw new ArgumentException("Наименование товара должно содержать от 2 до 80 символов.");
        }

        if (price < 0 || quantity < 0)
        {
            throw new ArgumentException("Цена и остаток товара не могут быть отрицательными.");
        }

        using var db = _factory.Create();

        var item = await db.Items
            .FirstOrDefaultAsync(entity => entity.Id_Item == itemId);

        if (item is null)
        {
            throw new InvalidOperationException($"Товар с идентификатором {itemId} не найден.");
        }

        item.Name_Item = name;
        item.NameArticule = string.IsNullOrWhiteSpace(articule) ? null : articule.Trim();
        item.Category_Id = categoryId;
        item.Price = price;
        item.Quantity = quantity;

        await db.SaveChangesAsync();
    }

    /// <summary>
    /// Удаляет товар, если по нему нет заказов.
    /// </summary>
    public async Task DeleteItemAsync(int itemId)
    {
        using var db = _factory.Create();

        var item = await db.Items
            .Include(entity => entity.Orders)
            .FirstOrDefaultAsync(entity => entity.Id_Item == itemId);

        if (item is null)
        {
            throw new InvalidOperationException($"Товар с идентификатором {itemId} не найден.");
        }

        if (item.Orders.Count > 0)
        {
            throw new InvalidOperationException(
                "Нельзя удалить товар, по которому оформлены заказы.");
        }

        db.Items.Remove(item);

        await db.SaveChangesAsync();
    }

    /// <summary>
    /// Добавляет категорию (отделение).
    /// </summary>
    public async Task AddCategoryAsync(string name)
    {
        name = name.Trim();

        if (name.Length is < 2 or > 50)
        {
            throw new ArgumentException("Название категории должно содержать от 2 до 50 символов.");
        }

        using var db = _factory.Create();

        var exists = await db.Categories
            .AnyAsync(category => category.Name_Category == name);

        if (exists)
        {
            throw new ArgumentException("Такая категория уже есть.");
        }

        db.Categories.Add(new Category { Name_Category = name });

        await db.SaveChangesAsync();
    }

    /// <summary>
    /// Удаляет категорию, если в ней нет товаров.
    /// </summary>
    public async Task DeleteCategoryAsync(int categoryId)
    {
        using var db = _factory.Create();

        var category = await db.Categories
            .Include(entity => entity.Items)
            .FirstOrDefaultAsync(entity => entity.Id_Category == categoryId);

        if (category is null)
        {
            throw new InvalidOperationException("Категория не найдена.");
        }

        if (category.Items.Count > 0)
        {
            throw new InvalidOperationException(
                "Нельзя удалить категорию, в которой есть товары.");
        }

        db.Categories.Remove(category);

        await db.SaveChangesAsync();
    }

    /// <summary>
    /// Список пользователей.
    /// </summary>
    public async Task<List<UserInfo>> LoadUsersAsync()
    {
        using var db = _factory.Create();

        var users = await db.Users
            .AsNoTracking()
            .Include(user => user.Role)
            .OrderBy(user => user.Login)
            .ToListAsync();

        return users
            .Select(user => new UserInfo
            {
                Id = user.Id_User,
                Login = user.Login,
                FullName = $"{user.FirstName} {user.SecondName}".Trim(),
                Role = user.Role?.Name_Role ?? "",
                RoleId = user.Role_Id
            })
            .ToList();
    }

    /// <summary>
    /// Меняет роль пользователя.
    /// </summary>
    public async Task ChangeUserRoleAsync(int userId, string roleName)
    {
        using var db = _factory.Create();

        var user = await db.Users
            .FirstOrDefaultAsync(entity => entity.Id_User == userId);

        if (user is null)
        {
            throw new InvalidOperationException("Пользователь не найден.");
        }

        var role = await db.Roles
            .FirstOrDefaultAsync(entity => entity.Name_Role == roleName);

        if (role is null)
        {
            throw new InvalidOperationException($"Роль «{roleName}» не найдена.");
        }

        user.Role_Id = role.Id_Role;

        await db.SaveChangesAsync();
    }

    /// <summary>
    /// Сводный отчёт по заказам.
    /// </summary>
    public async Task<ReportInfo> BuildReportAsync()
    {
        using var db = _factory.Create();

        var orders = await db.Orders
            .AsNoTracking()
            .Include(order => order.Item)
            .Include(order => order.Status)
            .ToListAsync();

        return new ReportInfo
        {
            OrdersCount = orders.Count,
            ItemsCount = await db.Items.CountAsync(),
            UsersCount = await db.Users.CountAsync(),
            ReviewsCount = await db.ItemReviews.CountAsync(),
            TotalAmount = orders.Sum(order => order.Item?.Price ?? 0m),
            CancelledCount = orders.Count(order => order.Status?.Name_Status == "Отменён"),
            ByStatus = orders
                .GroupBy(order => order.Status?.Name_Status ?? "Не указан")
                .Select(group => new ReportLine
                {
                    Title = group.Key,
                    Value = group.Count()
                })
                .OrderBy(line => line.Title)
                .ToList()
        };
    }
}

/// <summary>
/// Пользователь для интерфейса администратора.
/// </summary>
public sealed class UserInfo
{
    public int Id { get; set; }

    public string Login { get; set; } = "";

    public string FullName { get; set; } = "";

    public string Role { get; set; } = "";

    public int RoleId { get; set; }

    public override string ToString() => $"{FullName} ({Login}) — {Role}";
}

/// <summary>
/// Отчёт по заказам.
/// </summary>
public sealed class ReportInfo
{
    public int OrdersCount { get; set; }

    public int ItemsCount { get; set; }

    public int UsersCount { get; set; }

    public int ReviewsCount { get; set; }

    public int CancelledCount { get; set; }

    public decimal TotalAmount { get; set; }

    public List<ReportLine> ByStatus { get; set; } = new();
}

/// <summary>
/// Строка отчёта.
/// </summary>
public sealed class ReportLine
{
    public string Title { get; set; } = "";

    public int Value { get; set; }
}
