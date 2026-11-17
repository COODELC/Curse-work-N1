using Ktpt.CourseOrder.Data;
using Ktpt.CourseOrder.Data.Entities;
using Microsoft.AspNetCore.Identity;
using Microsoft.EntityFrameworkCore;
using System;
using System.Threading.Tasks;

namespace Ktpt.CourseOrder.Auth;

public sealed class AuthService
{
    public const int MinPasswordLength = 8;

    private readonly DbFactory _factory;

    private readonly PasswordHasher<AppUser> _hasher = new();

    public AuthService(DbFactory factory)
    {
        _factory = factory;
    }

    public async Task<AppUser?> LoginAsync(
        string login,
        string password)
    {
        login = login.Trim();

        if (login.Length == 0 || string.IsNullOrEmpty(password))
        {
            return null;
        }

        using var db = _factory.Create();

        var stored = await db.Users
            .Include(user => user.Role)
            .FirstOrDefaultAsync(user => user.Login == login);

        if (stored is null)
        {
            return null;
        }

        var role = ParseRole(stored.Role?.Name_Role ?? string.Empty);

        if (role is null)
        {
            return null;
        }

        var user = new AppUser(
            stored.Id_User,
            stored.Login,
            stored.FirstName,
            stored.SecondName,
            stored.Age,
            role.Value);

        PasswordVerificationResult result;

        try
        {
            result = _hasher.VerifyHashedPassword(
                user,
                stored.Password,
                password);
        }
        catch (FormatException)
        {
            return null;
        }

        if (result == PasswordVerificationResult.Failed)
        {
            return null;
        }

        if (result == PasswordVerificationResult.SuccessRehashNeeded)
        {
            stored.Password = _hasher.HashPassword(user, password);

            await db.SaveChangesAsync();
        }

        return user;
    }

    public async Task RegisterAsync(
        string firstName,
        string login,
        string password,
        string confirmation)
    {
        firstName = firstName.Trim();
        login = login.Trim();

        if (firstName.Length is < 1 or > 100)
        {
            throw new ArgumentException(
                "Имя должно содержать от 1 до 100 символов.");
        }

        if (login.Length is < 3 or > 100)
        {
            throw new ArgumentException(
                "Логин должен содержать от 3 до 100 символов.");
        }

        if (password.Length is < MinPasswordLength or > 128)
        {
            throw new ArgumentException(
                $"Пароль должен содержать от {MinPasswordLength} до 128 символов.");
        }

        if (password != confirmation)
        {
            throw new ArgumentException(
                "Пароли не совпадают.");
        }

        using var db = _factory.Create();

        var role = await db.Roles
            .FirstOrDefaultAsync(item => item.Name_Role == "Покупатель");

        if (role is null)
        {
            throw new InvalidOperationException(
                "В таблице Roles должна быть одна роль «Покупатель».");
        }

        var loginTaken = await db.Users
            .AnyAsync(user => user.Login == login);

        if (loginTaken)
        {
            throw new ArgumentException(
                "Этот логин уже занят.");
        }

        var user = new AppUser(
            0,
            login,
            firstName,
            null,
            null,
            UserRole.Customer);

        var entity = new Users
        {
            FirstName = firstName,
            Login = login,
            Password = _hasher.HashPassword(user, password),
            Role_Id = role.Id_Role
        };

        db.Users.Add(entity);

        await db.SaveChangesAsync();
    }

    public async Task<string> ChangeUserRoleAsync(
        int userId,
        int roleId,
        UserRole currentRole)
    {
        if (currentRole != UserRole.Administrator)
        {
            return "Изменять роли может только администратор.";
        }

        using var db = _factory.Create();

        var user = await db.Users
            .FirstOrDefaultAsync(item => item.Id_User == userId);

        if (user is null)
        {
            return "Пользователь не найден.";
        }

        var role = await db.Roles
            .FirstOrDefaultAsync(item => item.Id_Role == roleId);

        if (role is null)
        {
            return "Выбранная роль не найдена.";
        }

        user.Role_Id = role.Id_Role;

        await db.SaveChangesAsync();

        return $"Роль пользователя изменена на «{role.Name_Role}».";
    }

    public void EnsureDemoUsers()
    {
        using var db = _factory.Create();

        var admin = db.Roles.FirstOrDefault(role => role.Name_Role == "Администратор");
        var customer = db.Roles.FirstOrDefault(role => role.Name_Role == "Покупатель");
        var manager = db.Roles.FirstOrDefault(role => role.Name_Role == "Менеджер");

        if (admin is null || customer is null || manager is null)
        {
            return;
        }

        var accounts = new[]
        {
            new DemoAccount("Андрей", "Копейкин", 18, "ander", "8942", admin.Id_Role),
            new DemoAccount("Мария", "Иванова", 25, "maria", "1234", customer.Id_Role),
            new DemoAccount("Иван", "Петров", 30, "ivan", "5678", manager.Id_Role),
            new DemoAccount("Ольга", "Смирнова", 22, "olga", "abcd", customer.Id_Role),
            new DemoAccount("Сергей", "Волков", 41, "sergey", "pass1", customer.Id_Role)
        };

        foreach (var account in accounts)
        {
            var user = db.Users
                .FirstOrDefault(entity => entity.Login == account.Login);

            if (user is null)
            {
                db.Users.Add(CreateUser(account));

                continue;
            }

            if (!IsHashed(user.Password))
            {
                user.Password = _hasher.HashPassword(
                    new AppUser(
                        user.Id_User,
                        user.Login,
                        user.FirstName,
                        user.SecondName,
                        user.Age,
                        UserRole.Customer),
                    account.Password);
            }
        }

        db.SaveChanges();
    }

    private static bool IsHashed(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        try
        {
            var decoded = Convert.FromBase64String(value);

            return decoded.Length > 1
                && decoded[0] == 0x00
                && decoded[1] == 0x01;
        }
        catch (FormatException)
        {
            return false;
        }
    }

    private Users CreateUser(DemoAccount account)
    {
        var user = new AppUser(
            0,
            account.Login,
            account.FirstName,
            account.SecondName,
            account.Age,
            UserRole.Customer);

        return new Users
        {
            FirstName = account.FirstName,
            SecondName = account.SecondName,
            Age = account.Age,
            Login = account.Login,
            Password = _hasher.HashPassword(user, account.Password),
            Role_Id = account.RoleId
        };
    }

    private sealed record DemoAccount(
        string FirstName,
        string SecondName,
        int Age,
        string Login,
        string Password,
        int RoleId);

    private static UserRole? ParseRole(string name)
    {
        return name.Trim() switch
        {
            "Гость" => UserRole.Guest,
            "Покупатель" => UserRole.Customer,
            "Менеджер" => UserRole.Manager,
            "Администратор" => UserRole.Administrator,
            _ => null
        };
    }
}
