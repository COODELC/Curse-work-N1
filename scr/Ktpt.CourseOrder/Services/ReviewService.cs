using Ktpt.CourseOrder.Data;
using Ktpt.CourseOrder.Data.Entities;
using Microsoft.EntityFrameworkCore;

namespace Ktpt.CourseOrder.Services;

/// <summary>
/// Работа с отзывами о товарах.
/// </summary>
public sealed class ReviewService
{
    private readonly DbFactory _factory;

    public ReviewService(DbFactory factory)
    {
        _factory = factory;
    }

    /// <summary>
    /// Добавляет отзыв о товаре от имени пользователя.
    /// </summary>
    /// <param name="itemId">Товар, к которому относится отзыв.</param>
    /// <param name="text">Текст отзыва.</param>
    /// <param name="userId">Автор отзыва. null — гость.</param>
    public async Task AddReviewAsync(int itemId, string text, int? userId = null)
    {
        text = text.Trim();

        if (text.Length is < 3 or > 1000)
        {
            throw new ArgumentException(
                "Текст отзыва должен содержать от 3 до 1000 символов.");
        }

        using var db = _factory.Create();

        var itemExists = await db.Items
            .AnyAsync(item => item.Id_Item == itemId);

        if (!itemExists)
        {
            throw new InvalidOperationException(
                $"Товар с идентификатором {itemId} не найден в каталоге.");
        }

        if (userId is not null)
        {
            var userExists = await db.Users
                .AnyAsync(user => user.Id_User == userId.Value);

            if (!userExists)
            {
                throw new InvalidOperationException(
                    $"Пользователь с идентификатором {userId} не найден.");
            }
        }

        db.ItemReviews.Add(new ItemReview
        {
            Review = text,
            Item_Id = itemId,
            User_Id = userId
        });

        await db.SaveChangesAsync();
    }

    /// <summary>
    /// Отзывы по товару вместе с авторами.
    /// </summary>
    public async Task<List<Ktpt.CourseOrder.Models.ProductReview>> GetReviewsAsync(int itemId)
    {
        using var db = _factory.Create();

        var reviews = await db.ItemReviews
            .AsNoTracking()
            .Include(review => review.User)
            .ThenInclude(user => user!.Role)
            .Where(review => review.Item_Id == itemId)
            .OrderBy(review => review.Id_Review)
            .ToListAsync();

        return reviews
            .Select(review => new Ktpt.CourseOrder.Models.ProductReview
            {
                Id = review.Id_Review,
                Text = review.Review,
                Item_Id = review.Item_Id,
                Author = AuthorName(review),
                AuthorRole = review.User?.Role?.Name_Role ?? ""
            })
            .ToList();
    }

    /// <summary>
    /// Имя автора отзыва. Для отзыва без пользователя показывается «Гость».
    /// </summary>
    private static string AuthorName(ItemReview review)
    {
        var user = review.User;

        if (user is null)
        {
            return "Гость";
        }

        var name = $"{user.FirstName} {user.SecondName}".Trim();

        return name.Length > 0 ? name : user.Login;
    }
}
