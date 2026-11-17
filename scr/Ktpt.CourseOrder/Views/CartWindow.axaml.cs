using Avalonia.Controls;
using Ktpt.CourseOrder.ViewModels;

namespace Ktpt.CourseOrder.Views;

public partial class CartWindow : Window
{
    public CartWindow(MainViewModel mainViewModel)
    {
        InitializeComponent();

        DataContext = new CartViewModel(mainViewModel);
    }
}
