using Ktpt.CourseOrder.Auth;
using Ktpt.CourseOrder.Data;
using System.IO;

namespace Ktpt.CourseOrder.Models;

/// <summary>
/// Общие сервисы приложения: подключение к базе, авторизация, сеанс пользователя.
/// </summary>
public static class AppServices
{
    /// <summary>Фабрика контекста данных. Строка подключения — в Connection.txt.</summary>
    public static DbFactory Data { get; } = DbFactory.CreateFromFile();

    /// <summary>Текущий пользователь и его права.</summary>
    public static AuthSession Session { get; } = new();

    /// <summary>Сервис авторизации.</summary>
    public static AuthService Authentication { get; } = new(Data);
}
