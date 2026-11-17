using Avalonia.Controls;
using Ktpt.CourseOrder.ViewModels;

namespace Ktpt.CourseOrder.Views;

/// <summary>
/// Окно управления каталогом и пользователями для администратора.
/// </summary>
public partial class AdminWindow : Window
{
    public AdminWindow(MainViewModel mainViewModel)
    {
        InitializeComponent();

        DataContext = new AdminViewModel(mainViewModel);
    }
}
