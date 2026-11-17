using CommunityToolkit.Mvvm.ComponentModel;

namespace Ktpt.CourseOrder.Models;

/// <summary>
/// Позиция корзины. Наблюдаемый объект, чтобы количество и сумма
/// сразу отражались в окне корзины.
/// </summary>
public partial class CartItem : ObservableObject
{
    public int ProductId { get; set; }

    public string Name_Item { get; set; } = "";

    public decimal Price { get; set; }

    [ObservableProperty]
    private int quantity;

    public decimal Total => Price * Quantity;

    partial void OnQuantityChanged(int value) => OnPropertyChanged(nameof(Total));
}
