using System;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;

namespace Ktpt.CourseOrder.Auth;

public sealed class AuthWindow : Window
{
    private readonly AuthService _service;

    private readonly TextBlock _message = new()
    {
        TextWrapping = Avalonia.Media.TextWrapping.Wrap
    };

    public AuthWindow(AuthService service)
    {
        _service = service;

        Title = "Авторизация пользователя";
        Width = 460;
        Height = 580;
        CanResize = false;
        WindowStartupLocation =
            WindowStartupLocation.CenterOwner;

        var tabs = new TabControl
        {
            ItemsSource = new[]
            {
                new TabItem
                {
                    Header = "Вход",
                    Content = BuildLogin()
                },
                new TabItem
                {
                    Header = "Регистрация",
                    Content = BuildRegistration()
                }
            }
        };

        var panel = new StackPanel
        {
            Margin = new Thickness(24),
            Spacing = 16
        };

        panel.Children.Add(new TextBlock
        {
            Text = "Магазин автозапчастей",
            FontSize = 26
        });

        panel.Children.Add(tabs);
        panel.Children.Add(_message);

        Content = panel;
    }

    private Control BuildLogin()
    {
        var login = new TextBox
        {
            PlaceholderText = "Логин"
        };

        var password = new TextBox
        {
            PlaceholderText = "Пароль",
            PasswordChar = '●'
        };

        var button = new Button
        {
            Content = "Войти",
            HorizontalAlignment = HorizontalAlignment.Stretch
        };

        button.Click += async (_, _) =>
        {
            await RunAsync(button, async () =>
            {
                var user = await _service.LoginAsync(
                    login.Text ?? "",
                    password.Text ?? "");

                password.Text = "";

                if (user is null)
                {
                    _message.Text =
                        "Неверный логин или пароль.";

                    return;
                }

                Close(user);
            });
        };

        var panel = CreatePanel();
        panel.Children.Add(login);
        panel.Children.Add(password);
        panel.Children.Add(button);

        return panel;
    }

    private Control BuildRegistration()
    {
        var name = new TextBox { PlaceholderText = "Имя" };
        var login = new TextBox { PlaceholderText = "Логин" };

        var password = new TextBox
        {
            PlaceholderText =
                $"Пароль — от {AuthService.MinPasswordLength} символов",
            PasswordChar = '●'
        };

        var confirmation = new TextBox
        {
            PlaceholderText = "Повторите пароль",
            PasswordChar = '●'
        };

        var button = new Button
        {
            Content = "Зарегистрироваться",
            HorizontalAlignment = HorizontalAlignment.Stretch
        };

        button.Click += async (_, _) =>
        {
            await RunAsync(button, async () =>
            {
                await _service.RegisterAsync(
                    name.Text ?? "",
                    login.Text ?? "",
                    password.Text ?? "",
                    confirmation.Text ?? "");

                password.Text = "";
                confirmation.Text = "";

                _message.Text =
                    "Регистрация завершена. Перейдите на вкладку «Вход».";
            });
        };

        var panel = CreatePanel();
        panel.Children.Add(name);
        panel.Children.Add(login);
        panel.Children.Add(password);
        panel.Children.Add(confirmation);
        panel.Children.Add(button);

        return panel;
    }

    private async Task RunAsync(
        Button button,
        Func<Task> action)
    {
        button.IsEnabled = false;
        _message.Text = "";

        try
        {
            await action();
        }
        catch (ArgumentException ex)
        {
            _message.Text = ex.Message;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(ex);

            _message.Text =
                "Не удалось выполнить операцию. " +
                "Проверьте подключение к базе данных.";
        }
        finally
        {
            button.IsEnabled = true;
        }
    }

    private static StackPanel CreatePanel()
    {
        return new StackPanel
        {
            Spacing = 12,
            Margin = new Thickness(0, 16, 0, 0)
        };
    }
}
