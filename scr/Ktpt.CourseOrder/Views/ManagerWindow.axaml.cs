using Avalonia.Controls;
using Ktpt.CourseOrder.ViewModels;

namespace Ktpt.CourseOrder.Views;

/// <summary>
/// Окно обработки заказов для сотрудника и администратора.
/// </summary>
public partial class ManagerWindow : Window
{
    public ManagerWindow(MainViewModel mainViewModel)
    {
        InitializeComponent();

        DataContext = new ManagerViewModel(mainViewModel);
    }
}
