using System.Collections.Generic;

namespace Ktpt.CourseOrder.Models;

/// <summary>
/// Товар каталога, который отображается в интерфейсе.
/// </summary>
public class Product
{
    public int Id { get; set; }

    public string Name_Item { get; set; } = "";

    public string NameArticule { get; set; } = "";

    public int Category_Id { get; set; }

    public string Name_Category { get; set; } = "";

    public decimal Price { get; set; }

    public int Quantity { get; set; }

    public string ImagePath { get; set; } = "/Assets/Products/no-image.png";

    public string Description { get; set; } = "";

    public List<ProductReview> Reviews { get; set; } = new();

    public ProductReview? PreviewReview { get; set; }

    public string PriceText => Price.ToString("N2") + " руб.";

    public string QuantityText => Quantity > 0
        ? $"В наличии: {Quantity} шт."
        : "Нет в наличии";
}

/// <summary>
/// Отзыв о товаре.
/// </summary>
public class ProductReview
{
    public int Id { get; set; }

    public string Text { get; set; } = "";

    public int Item_Id { get; set; }

    public string Author { get; set; } = "Гость";

    public string AuthorRole { get; set; } = "";
}
