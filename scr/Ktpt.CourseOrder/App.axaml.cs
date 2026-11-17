using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Markup.Xaml;
using Ktpt.CourseOrder.Data;
using Ktpt.CourseOrder.Models;
using System;
using System.Diagnostics;

namespace Ktpt.CourseOrder;

public partial class App : Application
{
    public override void Initialize()
    {
        AvaloniaXamlLoader.Load(this);
    }

    public override void OnFrameworkInitializationCompleted()
    {
        InitializeDatabase();

        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
        {
            desktop.MainWindow = new Views.MainWindow();
        }

        base.OnFrameworkInitializationCompleted();
    }

    /// <summary>
    /// Создаёт структуру таблиц и заполняет базу начальными данными.
    /// Основной режим — SQL Server. Если сервер недоступен,
    /// приложение переключается на локальную базу данных, чтобы
    /// программа оставалась работоспособной.
    /// </summary>
    private static void InitializeDatabase()
    {
        try
        {
            EnsureDatabase(AppServices.Data);
        }
        catch (Exception ex)
        {
            Debug.WriteLine(ex);

            AppServices.Data.SwitchToLocalFile();

            EnsureDatabase(AppServices.Data);
        }
    }

    private static void EnsureDatabase(DbFactory factory)
    {
        using var db = factory.Create();

        db.Database.EnsureCreated();

        // Новые колонки в уже созданную базу EnsureCreated() не добавляет.
        SchemaUpgrade.Apply(db);

        // Роли, категории, товары, статусы и отзывы.
        SeedData.Apply(db);

        // Пользователи: хеши паролей вычисляет модуль авторизации.
        AppServices.Authentication.EnsureDemoUsers();

        // Заказы появляются после создания пользователей.
        SeedData.Apply(db);

        // Авторы отзывов проставляются после появления пользователей и заказов.
        SeedData.AssignReviewAuthors(db);

        // Описания товаров заполняются, если колонка появилась миграцией.
        SeedData.AssignItemDescriptions(db);
    }
}
