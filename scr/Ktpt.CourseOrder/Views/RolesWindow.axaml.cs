using Avalonia.Controls;
using Avalonia.Interactivity;
using Ktpt.CourseOrder.ViewModels;

namespace Ktpt.CourseOrder.Views;

public partial class RolesWindow : Window
{
    public RolesWindow()
    {
        InitializeComponent();

        DataContext = new RolesViewModel();
    }
}
