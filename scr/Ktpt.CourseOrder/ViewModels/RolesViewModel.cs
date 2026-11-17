using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Ktpt.CourseOrder.Auth;
using Ktpt.CourseOrder.Data;
using Ktpt.CourseOrder.Data.Entities;
using Ktpt.CourseOrder.Models;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;

namespace Ktpt.CourseOrder.ViewModels;

public partial class RolesViewModel : ObservableObject
{
    private readonly DbFactory _factory;

    public ObservableCollection<RoleOption> Roles { get; } = new();

    public ObservableCollection<UserRoleItem> Users { get; } = new();

    [ObservableProperty]
    private string statusText = "Загрузка пользователей...";

    public RolesViewModel()
    {
        _factory = AppServices.Data;

        _ = LoadAsync();
    }

    private async Task LoadAsync()
    {
        try
        {
            using var db = _factory.Create();

            var roles = await db.Roles
                .AsNoTracking()
                .OrderBy(role => role.Id_Role)
                .ToListAsync();

            var users = await db.Users
                .AsNoTracking()
                .Include(user => user.Role)
                .OrderBy(user => user.SecondName ?? string.Empty)
                .ThenBy(user => user.FirstName)
                .ToListAsync();

            Roles.Clear();

            foreach (var role in roles)
            {
                Roles.Add(new RoleOption(role.Id_Role, role.Name_Role));
            }

            Users.Clear();

            foreach (var user in users)
            {
                var selected = Roles.FirstOrDefault(
                    item => item.Id == user.Role_Id);

                Users.Add(new UserRoleItem
                {
                    Id = user.Id_User,
                    DisplayName = $"{user.FirstName} {user.SecondName}".Trim(),
                    Login = user.Login,
                    SelectedRole = selected
                });
            }

            StatusText = $"Пользователей: {Users.Count}";
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(ex);

            StatusText = "Не удалось загрузить пользователей.";
        }
    }

    [RelayCommand]
    private async Task SaveRoleAsync(UserRoleItem? item)
    {
        if (item is null || item.SelectedRole is null)
        {
            return;
        }

        try
        {
            using var db = _factory.Create();

            var user = await db.Users
                .FirstOrDefaultAsync(entity => entity.Id_User == item.Id);

            if (user is null)
            {
                StatusText = "Пользователь не найден.";

                return;
            }

            user.Role_Id = item.SelectedRole.Id;

            await db.SaveChangesAsync();

            StatusText =
                $"Роль пользователя {item.DisplayName} изменена " +
                $"на «{item.SelectedRole.Name}».";
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(ex);

            StatusText = "Не удалось изменить роль пользователя.";
        }
    }
}

public sealed record RoleOption(int Id, string Name)
{
    public override string ToString() => Name;
}

/// <summary>
/// Пользователь, роль которого редактируется.
/// Свойства объявлены с обычными установщиками: запись позволяет
/// двухсторонней привязке сохранить выбранную роль в ComboBox.
/// </summary>
public sealed class UserRoleItem
{
    public int Id { get; set; }

    public string DisplayName { get; set; } = "";

    public string Login { get; set; } = "";

    public RoleOption? SelectedRole { get; set; }
}
