using Ktpt.CourseOrder.Data.Entities;
using Microsoft.EntityFrameworkCore;

namespace Ktpt.CourseOrder.Data;

/// <summary>
/// Начальное наполнение базы данных.
/// Пользователей заполняет модуль авторизации: там же вычисляются хеши паролей.
/// </summary>
public static class SeedData
{
    public static void Apply(WorkshopDbContext context)
    {
        SeedRoles(context);
        SeedCategories(context);
        SeedItems(context);
        SeedOrderStatuses(context);
        SeedReviews(context);
        SeedOrders(context);
    }

    /// <summary>
    /// Описания товаров каталога. Используются при заполнении
    /// новой базы и при обновлении уже созданной.
    /// </summary>
    private static readonly Dictionary<string, string> ItemDescriptions = new()
    {
        ["Масляный насос"] =
            "Масляный насос двигателя для подачи смазки под давлением к деталям силовой установки. Деталь износостойкая, подходит для стандартных систем смазки.",
        ["Воздушный фильтр"] =
            "Воздушный фильтр для очистки поступающего в двигатель воздуха от пыли и механических частиц. Фильтрующий элемент меняется по мере засорения.",
        ["Тормозные колодки"] =
            "Комплект тормозных колодок для замены в тормозной системе автомобиля. Отличаются устойчивым сцеплением и низким уровнем шума при торможении.",
        ["Тормозной диск"] =
            "Тормозной диск (ротор) колёсного тормозного механизма. Вентилируемая конструкция снижает нагрев и обеспечивает равномерное торможение.",
        ["Амортизатор"] =
            "Амортизатор в сборе для подвески автомобиля. Гасит колебания кузова и обеспечивает плавный ход на неровностях дорожного покрытия.",
        ["Шаровая опора"] =
            "Шаровая опора для направляющего аппарата подвески. Обеспечивает устойчивое соединение ступицы с рычагом подвески.",
        ["Аккумулятор"] =
            "Аккумуляторная батарея для стартерного и осветительного оборудования. Обслуживаемая, с низким уровнем саморазряда.",
        ["Свеча зажигания"] =
            "Свеча зажигания для двигателя. Обеспечивает надёжное воспламенение топливно-воздушной смеси в цилиндре.",
        ["Моторное масло"] =
            "Моторное масло для бензиновых двигателей. Выполняет смазывающую, охлаждающую, очищающую и антиокислительную функции.",
        ["Антифриз"] =
            "Охлаждающая жидкость для системы охлаждения двигателя. Защищает систему от замерзания и перегрева в любой сезон.",
        ["Бампер"] =
            "Бампер автомобиля с заводским оформлением. Устанавливается на штатные крепления кузова без доработок.",
        ["Крыло"] =
            "Крыло автомобиля — элемент кузовного оперения. Подходит для восстановительного ремонта и покраски по стандарту."
    };

    /// <summary>
    /// Заполняет пустые описания товаров.
    /// </summary>
    public static void AssignItemDescriptions(WorkshopDbContext context)
    {
        var items = context.Items.ToList();

        var changed = false;

        foreach (var item in items)
        {
            if (!string.IsNullOrWhiteSpace(item.Description))
            {
                continue;
            }

            if (!ItemDescriptions.TryGetValue(item.Name_Item, out var description))
            {
                continue;
            }

            item.Description = description;
            changed = true;
        }

        if (changed)
        {
            context.SaveChanges();
        }
    }

    /// <summary>
    /// Проставляет авторов отзывам, у которых он не указан.
    /// Автором считается покупатель, который заказывал этот товар;
    /// если таких покупателей нет — любой пользователь с ролью «Покупатель».
    /// </summary>
    public static void AssignReviewAuthors(WorkshopDbContext context)
    {

        var orphanReviews = context.ItemReviews
            .Where(review => review.User_Id == null)
            .ToList();

        if (orphanReviews.Count == 0)
        {
            return;
        }

        var customers = context.Users
            .Where(user => user.Role != null && user.Role.Name_Role == "Покупатель")
            .OrderBy(user => user.Id_User)
            .Select(user => user.Id_User)
            .ToList();

        if (customers.Count == 0)
        {
            return;
        }

        var buyerByItem = context.Orders
            .AsNoTracking()
            .GroupBy(order => order.Item_Id)
            .Select(group => new
            {
                ItemId = group.Key,
                UserId = group
                    .OrderByDescending(order => order.Date_Order)
                    .ThenByDescending(order => order.Id_Order)
                    .Select(order => order.UserId)
                    .First()
            })
            .ToDictionary(item => item.ItemId, item => item.UserId);

        var fallbackIndex = 0;

        foreach (var review in orphanReviews)
        {
            int? authorId = null;

            if (buyerByItem.TryGetValue(review.Item_Id, out var buyerId)
                && context.Users.Any(user => user.Id_User == buyerId))
            {
                authorId = buyerId;
            }
            else
            {
                authorId = customers[fallbackIndex % customers.Count];
                fallbackIndex++;
            }

            review.User_Id = authorId;
        }

        context.SaveChanges();
    }

    private static void SeedRoles(WorkshopDbContext context)
    {
        if (context.Roles.Any())
        {
            return;
        }

        context.Roles.AddRange(
            new Roles { Name_Role = "Гость" },
            new Roles { Name_Role = "Покупатель" },
            new Roles { Name_Role = "Менеджер" },
            new Roles { Name_Role = "Администратор" });

        context.SaveChanges();
    }

    private static void SeedCategories(WorkshopDbContext context)
    {
        if (context.Categories.Any())
        {
            return;
        }

        context.Categories.AddRange(
            new Category { Name_Category = "Двигатель" },
            new Category { Name_Category = "Тормозная система" },
            new Category { Name_Category = "Подвеска" },
            new Category { Name_Category = "Электрика" },
            new Category { Name_Category = "Масла и жидкости" },
            new Category { Name_Category = "Кузовные детали" });

        context.SaveChanges();
    }

    private static void SeedItems(WorkshopDbContext context)
    {
        if (context.Items.Any())
        {
            return;
        }

        var items = new[]
        {
            new Item { Name_Item = "Масляный насос",      NameArticule = "Помпа",                    Category_Id = 1, Price = 4500m, Quantity = 10, ImagePath = "/Assets/Products/oil-pump.png" },
            new Item { Name_Item = "Воздушный фильтр",     NameArticule = "Фильтр очистки воздуха",  Category_Id = 1, Price = 850m,  Quantity = 25, ImagePath = "/Assets/Products/air-filter.png" },
            new Item { Name_Item = "Тормозные колодки",    NameArticule = "Колодки",                  Category_Id = 2, Price = 2300m, Quantity = 18, ImagePath = "/Assets/Products/brake-pads.png" },
            new Item { Name_Item = "Тормозной диск",       NameArticule = "Тормозной ротор",         Category_Id = 2, Price = 3200m, Quantity = 14, ImagePath = "/Assets/Products/brake-disc.png" },
            new Item { Name_Item = "Амортизатор",          NameArticule = "Стойка",                   Category_Id = 3, Price = 5600m, Quantity = 8,  ImagePath = "/Assets/Products/shock-absorber.png" },
            new Item { Name_Item = "Шаровая опора",        NameArticule = "Опора шаровая",            Category_Id = 3, Price = 1750m, Quantity = 20, ImagePath = "/Assets/Products/ball-joint.png" },
            new Item { Name_Item = "Аккумулятор",          NameArticule = "АКБ",                      Category_Id = 4, Price = 7800m, Quantity = 6,  ImagePath = "/Assets/Products/battery.png" },
            new Item { Name_Item = "Свеча зажигания",      NameArticule = "Свеча",                    Category_Id = 4, Price = 420m,  Quantity = 40, ImagePath = "/Assets/Products/spark-plug.png" },
            new Item { Name_Item = "Моторное масло",       NameArticule = "Машинное масло",           Category_Id = 5, Price = 3100m, Quantity = 30, ImagePath = "/Assets/Products/engine-oil.png" },
            new Item { Name_Item = "Антифриз",             NameArticule = "Охлаждающая жидкость",     Category_Id = 5, Price = 1250m, Quantity = 22, ImagePath = "/Assets/Products/antifreeze.png" },
            new Item { Name_Item = "Бампер",               NameArticule = "Буфер",                    Category_Id = 6, Price = 9400m, Quantity = 4,  ImagePath = "/Assets/Products/bumper.png" },
            new Item { Name_Item = "Крыло",                NameArticule = "Оперение",                 Category_Id = 6, Price = 6700m, Quantity = 5,  ImagePath = "/Assets/Products/fender.png" }
        };

        context.Items.AddRange(items);

        foreach (var item in items)
        {
            if (ItemDescriptions.TryGetValue(item.Name_Item, out var description))
            {
                item.Description = description;
            }
        }

        context.SaveChanges();
    }

    private static void SeedOrderStatuses(WorkshopDbContext context)
    {
        if (context.OrderStatuses.Any())
        {
            return;
        }

        context.OrderStatuses.AddRange(
            new OrderStatus { Name_Status = "Оформление заказа" },
            new OrderStatus { Name_Status = "Подтверждён" },
            new OrderStatus { Name_Status = "Собирается" },
            new OrderStatus { Name_Status = "Передан в доставку" },
            new OrderStatus { Name_Status = "Доставлен" },
            new OrderStatus { Name_Status = "Отменён" });

        context.SaveChanges();
    }

    private static void SeedReviews(WorkshopDbContext context)
    {
        if (context.ItemReviews.Any())
        {
            return;
        }

        var reviews = new[]
        {
            new ItemReview { Review = "Отличный масляный насос, работает без нареканий.", Item_Id = 1 },
            new ItemReview { Review = "Фильтр подошёл идеально, установка простая.",        Item_Id = 2 },
            new ItemReview { Review = "Колодки тормозят уверенно, не скрипят.",            Item_Id = 3 },
            new ItemReview { Review = "Диск качественный, покрытие ровное.",              Item_Id = 4 },
            new ItemReview { Review = "Амортизатор заметно улучшил ход машины.",          Item_Id = 5 },
            new ItemReview { Review = "Шаровая опора установилась без проблем.",          Item_Id = 6 },
            new ItemReview { Review = "Аккумулятор держит заряд хорошо.",                  Item_Id = 7 },
            new ItemReview { Review = "Свечи работают стабильно.",                        Item_Id = 8 },
            new ItemReview { Review = "Масло подходящее, двигатель работает тише.",        Item_Id = 9 },
            new ItemReview { Review = "Антифриз не замерзает, цвет соответствует описанию.", Item_Id = 10 },
            new ItemReview { Review = "Бампер встал точно по креплениям.",                 Item_Id = 11 },
            new ItemReview { Review = "Крыло качественное, без дефектов.",                Item_Id = 12 },
            new ItemReview { Review = "Хорошее соотношение цены и качества.",             Item_Id = 3 },
            new ItemReview { Review = "Быстрая доставка, товар оригинальный.",             Item_Id = 7 },
            new ItemReview { Review = "Рекомендую, заменил без обращения в сервис.",       Item_Id = 5 }
        };

        context.ItemReviews.AddRange(reviews);

        context.SaveChanges();
    }

    private static void SeedOrders(WorkshopDbContext context)
    {
        if (context.Orders.Any())
        {
            return;
        }

        var maria = context.Users.FirstOrDefault(user => user.Login == "maria");
        var ivan = context.Users.FirstOrDefault(user => user.Login == "ivan");
        var olga = context.Users.FirstOrDefault(user => user.Login == "olga");
        var sergey = context.Users.FirstOrDefault(user => user.Login == "sergey");

        if (maria is null || ivan is null || olga is null || sergey is null)
        {
            // Пользователи ещё не заведены — заполним их отдельно.
            return;
        }

        context.Orders.AddRange(
            new Orders { Date_Order = new DateTime(2026, 10, 1),  Item_Id = 1,  UserId = maria.Id_User,   Status_Id = 5 },
            new Orders { Date_Order = new DateTime(2026, 10, 3),  Item_Id = 3,  UserId = maria.Id_User,   Status_Id = 5 },
            new Orders { Date_Order = new DateTime(2026, 10, 5),  Item_Id = 5,  UserId = ivan.Id_User,    Status_Id = 4 },
            new Orders { Date_Order = new DateTime(2026, 10, 8),  Item_Id = 7,  UserId = olga.Id_User,    Status_Id = 4 },
            new Orders { Date_Order = new DateTime(2026, 10, 12), Item_Id = 9,  UserId = sergey.Id_User,  Status_Id = 3 },
            new Orders { Date_Order = new DateTime(2026, 10, 15), Item_Id = 2,  UserId = maria.Id_User,   Status_Id = 3 },
            new Orders { Date_Order = new DateTime(2026, 10, 18), Item_Id = 4,  UserId = ivan.Id_User,    Status_Id = 2 },
            new Orders { Date_Order = new DateTime(2026, 10, 21), Item_Id = 6,  UserId = olga.Id_User,    Status_Id = 2 },
            new Orders { Date_Order = new DateTime(2026, 10, 25), Item_Id = 8,  UserId = sergey.Id_User,  Status_Id = 2 },
            new Orders { Date_Order = new DateTime(2026, 10, 28), Item_Id = 10, UserId = maria.Id_User,   Status_Id = 1 },
            new Orders { Date_Order = new DateTime(2026, 10, 29), Item_Id = 11, UserId = ivan.Id_User,    Status_Id = 1 },
            new Orders { Date_Order = new DateTime(2026, 10, 30), Item_Id = 12, UserId = olga.Id_User,    Status_Id = 6 });

        context.SaveChanges();
    }
}
