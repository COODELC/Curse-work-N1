using Avalonia.Controls;
using Avalonia.Interactivity;
using Ktpt.CourseOrder.Models;

namespace Ktpt.CourseOrder.Views;

public partial class ProductDetailsWindow : Window
{
    public ProductDetailsWindow(Product product)
    {
        InitializeComponent();

        DataContext = product;

        // Если описание не заполнено, показываются фактические данные товара.
        DescriptionText.Text = string.IsNullOrWhiteSpace(product.Description)
            ? $"«{product.Name_Item}» — товар категории «{product.Name_Category}». " +
              $"Артикул: {product.NameArticule}. Цена: {product.Price:N2} руб. " +
              $"Остаток на складе: {product.Quantity} шт."
            : product.Description;
    }

    private void Close_Click(object? sender, RoutedEventArgs e)
    {
        Close();
    }
}
