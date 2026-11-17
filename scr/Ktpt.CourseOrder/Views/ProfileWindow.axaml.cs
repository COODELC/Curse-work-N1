using Avalonia.Controls;
using Avalonia.Interactivity;
using Ktpt.CourseOrder.Auth;
using System;

namespace Ktpt.CourseOrder.Views;

public partial class ProfileWindow : Window
{
    public ProfileWindow(AppUser user)
    {
        InitializeComponent();

        DataContext = new ProfileData(
            user.FirstName,
            user.SecondName ?? "Не указана",
            user.Age is null ? "Не указан" : $"{user.Age} лет",
            user.Login,
            GetRoleName(user.Role));
    }

    private void Close_Click(object? sender, RoutedEventArgs e)
    {
        Close();
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

    public sealed record ProfileData(
        string FirstName,
        string SecondName,
        string AgeText,
        string Login,
        string RoleName);
}
