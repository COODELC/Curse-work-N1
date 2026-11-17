namespace Ktpt.CourseOrder.Data.Entities;

/// <summary>
/// Роль пользователя: Гость, Покупатель, Менеджер, Администратор.
/// </summary>
public class Roles
{
    public int Id_Role { get; set; }

    public string Name_Role { get; set; } = string.Empty;

    public ICollection<Users> Users { get; set; } = new List<Users>();
}

/// <summary>
/// Пользователь системы. Пароль хранится только в виде хеша.
/// </summary>
public class Users
{
    public int Id_User { get; set; }

    public string FirstName { get; set; } = string.Empty;

    public string? SecondName { get; set; }

    public int? Age { get; set; }

    public string Login { get; set; } = string.Empty;

    /// <summary>
    /// Пароль пользователя. В программе хранится только хеш пароля,
    /// открытый пароль в базе не сохраняется.
    /// </summary>
    public string Password { get; set; } = string.Empty;

    public int Role_Id { get; set; }

    public Roles Role { get; set; } = null!;

    public ICollection<Orders> Orders { get; set; } = new List<Orders>();

    /// <summary>Отзывы, оставленные пользователем.</summary>
    public ICollection<ItemReview> Reviews { get; set; } = new List<ItemReview>();
}

/// <summary>
/// Категория (отделение) автозапчастей.
/// </summary>
public class Category
{
    public int Id_Category { get; set; }

    public string Name_Category { get; set; } = string.Empty;

    public ICollection<Item> Items { get; set; } = new List<Item>();
}

/// <summary>
/// Товар — автозапчасть.
/// </summary>
public class Item
{
    public int Id_Item { get; set; }

    public string Name_Item { get; set; } = string.Empty;

    /// <summary>Альтернативное название (артикул).</summary>
    public string? NameArticule { get; set; }

    public int Category_Id { get; set; }

    public decimal Price { get; set; }

    /// <summary>Остаток на складе.</summary>
    public int Quantity { get; set; }

    /// <summary>Подробное описание товара.</summary>
    public string? Description { get; set; }

    public string? ImagePath { get; set; }

    public Category Category { get; set; } = null!;

    public ICollection<ItemReview> Reviews { get; set; } = new List<ItemReview>();

    public ICollection<Orders> Orders { get; set; } = new List<Orders>();
}

/// <summary>
/// Отзыв о товаре.
/// </summary>
public class ItemReview
{
    public int Id_Review { get; set; }

    public string Review { get; set; } = string.Empty;

    public int Item_Id { get; set; }

    /// <summary>
    /// Автор отзыва. Может быть пустым — отзыв без автора показывается как гостевой.
    /// </summary>
    public int? User_Id { get; set; }

    public Item Item { get; set; } = null!;

    public Users? User { get; set; }
}

/// <summary>
/// Статус заказа.
/// </summary>
public class OrderStatus
{
    public int Id_status { get; set; }

    public string Name_Status { get; set; } = string.Empty;

    public ICollection<Orders> Orders { get; set; } = new List<Orders>();
}

/// <summary>
/// Строка заказа: одна запчасть, заказанная пользователем.
/// </summary>
public class Orders
{
    public int Id_Order { get; set; }

    public DateTime Date_Order { get; set; }

    public int Item_Id { get; set; }

    public int UserId { get; set; }

    public int Status_Id { get; set; }

    public Item Item { get; set; } = null!;

    public Users User { get; set; } = null!;

    public OrderStatus Status { get; set; } = null!;
}
