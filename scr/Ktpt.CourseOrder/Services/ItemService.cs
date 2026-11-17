using Ktpt.CourseOrder.Data;
using Ktpt.CourseOrder.Models;
using Microsoft.EntityFrameworkCore;

namespace Ktpt.CourseOrder.Services;

/// <summary>
/// Работа с каталогом товаров: выборка, поиск, фильтрация по категориям.
/// </summary>
public sealed class ItemService
{
    private readonly DbFactory _factory;

    public ItemService(DbFactory factory)
    {
        _factory = factory;
    }

    /// <summary>
    /// Возвращает каталог товаров с названием категории и отзывами.
    /// </summary>
    public async Task<List<Product>> LoadItemsAsync(int? categoryId = null, string? search = null)
    {
        using var db = _factory.Create();

        var query = db.Items
            .AsNoTracking()
            .Include(item => item.Category)
            .Include(item => item.Reviews)
            .ThenInclude(review => review.User)
            .ThenInclude(user => user!.Role)
            .AsQueryable();

        if (categoryId is not null)
        {
            query = query.Where(item => item.Category_Id == categoryId.Value);
        }

        if (!string.IsNullOrWhiteSpace(search))
        {
            var pattern = $"%{search.Trim()}%";

            query = query.Where(item =>
                EF.Functions.Like(item.Name_Item, pattern)
                || (item.NameArticule != null
                    && EF.Functions.Like(item.NameArticule, pattern)));
        }

        var items = await query
            .OrderBy(item => item.Name_Item)
            .ToListAsync();

        return items.Select(Map).ToList();
    }

    private static Product Map(Ktpt.CourseOrder.Data.Entities.Item item)    {
        var reviews = item.Reviews
            .OrderBy(review => review.Id_Review)
            .Select(review => new ProductReview
            {
                Id = review.Id_Review,
                Text = review.Review,
                Item_Id = review.Item_Id,
                Author = review.User is null
                    ? "Гость"
                    : $"{review.User.FirstName} {review.User.SecondName}".Trim(),
                AuthorRole = review.User?.Role?.Name_Role ?? ""
            })
            .ToList();

        return new Product
        {
            Id = item.Id_Item,
            Name_Item = item.Name_Item,
            NameArticule = item.NameArticule ?? "",
            Category_Id = item.Category_Id,
            Name_Category = item.Category?.Name_Category ?? "",
            Price = item.Price,
            Quantity = item.Quantity,
            ImagePath = string.IsNullOrWhiteSpace(item.ImagePath)
                ? "/Assets/Products/no-image.png"
                : item.ImagePath,
            Description = item.Description ?? "",
            Reviews = reviews,

            // В карточке товара показывается только один отзыв.
            PreviewReview = reviews.FirstOrDefault()
        };
    }
}
