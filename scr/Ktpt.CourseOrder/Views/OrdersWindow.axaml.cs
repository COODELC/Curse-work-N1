using Avalonia.Controls;
using Ktpt.CourseOrder.ViewModels;

namespace Ktpt.CourseOrder.Views;

/// <summary>
/// Окно истории покупок пользователя.
/// </summary>
public partial class OrdersWindow : Window
{
    public OrdersWindow(MainViewModel mainViewModel)
    {
        InitializeComponent();

        DataContext = new OrdersViewModel(mainViewModel);
    }
}
