using Ktpt.CourseOrder.Data;
using Ktpt.CourseOrder.Models;
using Microsoft.EntityFrameworkCore;

namespace Ktpt.CourseOrder.Services;

/// <summary>
/// Работа с категориями (отделениями) товаров.
/// </summary>
public sealed class CategoryService
{
    private readonly DbFactory _factory;

    public CategoryService(DbFactory factory)
    {
        _factory = factory;
    }

    public async Task<List<CategoryOption>> LoadCategoriesAsync()
    {
        using var db = _factory.Create();

        var categories = await db.Categories
            .AsNoTracking()
            .OrderBy(category => category.Name_Category)
            .ToListAsync();

        return categories
            .Select(category => new CategoryOption
            {
                Id = category.Id_Category,
                Name = category.Name_Category
            })
            .ToList();
    }
}

/// <summary>
/// Категория товара в виде, пригодном для выбора в интерфейсе.
/// </summary>
public sealed class CategoryOption
{
    public int Id { get; set; }

    public string Name { get; set; } = "";

    public override string ToString() => Name;
}
