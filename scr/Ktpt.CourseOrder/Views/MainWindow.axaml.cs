using Avalonia.Controls;
using Avalonia.Interactivity;
using Ktpt.CourseOrder.Auth;
using Ktpt.CourseOrder.Models;
using Ktpt.CourseOrder.ViewModels;
using System;

namespace Ktpt.CourseOrder.Views;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();

        DataContext = new MainViewModel();

        if (DataContext is MainViewModel viewModel)
        {
            viewModel.Owner = this;
        }

        // Шапка обновляется автоматически при входе и выходе пользователя.
        AppServices.Session.Changed += () => UpdateAccountHeader();

        UpdateAccountHeader();
    }
    private async void ProfileButton_Click(
    object? sender,
    RoutedEventArgs e)
    {
        var user = AppServices.Session.CurrentUser;

        if (user is null)
        {
            return;
        }

        var window = new ProfileWindow(user);

        await window.ShowDialog(this);
    }

    private async void RolesButton_Click(
        object? sender,
        RoutedEventArgs e)
    {
        if (!AppServices.Session.Can(Permission.ManageUsers))
        {
            return;
        }

        var window = new RolesWindow();

        await window.ShowDialog(this);
    }
    private async void AccountButton_Click(
        object? sender,
        RoutedEventArgs e)
    {
        if (AppServices.Session.IsAuthenticated)
        {
            return;
        }

        AccountButton.IsEnabled = false;

        try
        {
            var window = new AuthWindow(
                AppServices.Authentication);

            var user = await window.ShowDialog<AppUser?>(this);

            // Пользователь закрыл окно без входа.
            if (user is null)
            {
                return;
            }

            AppServices.Session.SignIn(user);

            UpdateAccountHeader();
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(ex);

            if (DataContext is MainViewModel viewModel)
            {
                viewModel.StatusText =
                    "Не удалось выполнить вход. " +
                    "Проверьте настройки подключения.";
            }
        }
        finally
        {
            AccountButton.IsEnabled = true;
        }
    }

    private void LogoutButton_Click(
        object? sender,
        RoutedEventArgs e)
    {
        AppServices.Session.SignOut();

        if (DataContext is MainViewModel viewModel)
        {
            viewModel.HandleSignOut();
        }

        UpdateAccountHeader();
    }

    private void UpdateAccountHeader()
    {
        var user = AppServices.Session.CurrentUser;

        AccountButton.IsVisible = user is null;

        LogoutButton.IsVisible = user is not null;

        MyOrdersButton.IsVisible = user is not null;
        ProfileButton.IsVisible = user is not null;

        RolesButton.IsVisible =
            AppServices.Session.Can(Permission.ManageUsers);
        ManagerButton.IsVisible =
            AppServices.Session.Can(Permission.ManageOrders);

        AdminButton.IsVisible =
            AppServices.Session.Can(Permission.ManageProducts);

        CurrentUserText.Text = user is null
            ? "Гость"
            : $"{user.FirstName} — {GetRoleName(user.Role)}";
    }

    private static string GetRoleName(UserRole role)
    {
        return role switch
        {
            UserRole.Customer => "Покупатель",
            UserRole.Manager => "Менеджер",
            UserRole.Administrator => "Администратор",
            _ => "Гость"
        };
    }
}
