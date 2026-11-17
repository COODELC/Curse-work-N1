using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Ktpt.CourseOrder.Models;
using Ktpt.CourseOrder.Services;
using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;

namespace Ktpt.CourseOrder.ViewModels;

/// <summary>
/// Управление каталогом и пользователями — форма администратора.
/// </summary>
public partial class AdminViewModel : ObservableObject
{
    private readonly MainViewModel _mainViewModel;

    private readonly CatalogAdminService _adminService;

    private readonly ItemService _itemService;

    private readonly CategoryService _categoryService;

    public AdminViewModel(MainViewModel mainViewModel)
    {
        _mainViewModel = mainViewModel;

        _adminService = new CatalogAdminService(AppServices.Data);
        _itemService = new ItemService(AppServices.Data);
        _categoryService = new CategoryService(AppServices.Data);

        _ = LoadAsync();
    }

    public ObservableCollection<Product> Items { get; } = new();

    public ObservableCollection<CategoryOption> Categories { get; } = new();

    public ObservableCollection<UserInfo> Users { get; } = new();

    public ObservableCollection<string> Roles { get; } = new();

    public ObservableCollection<ReportLine> ReportLines { get; } = new();

    [ObservableProperty]
    private Product? selectedItem;

    [ObservableProperty]
    private UserInfo? selectedUser;

    [ObservableProperty]
    private string selectedRole = "Покупатель";

    [ObservableProperty]
    private CategoryOption? selectedItemCategory;

    [ObservableProperty]
    private string itemName = "";

    [ObservableProperty]
    private string itemArticule = "";

    [ObservableProperty]
    private string itemPrice = "0";

    [ObservableProperty]
    private string itemQuantity = "0";

    [ObservableProperty]
    private string newCategoryName = "";

    [ObservableProperty]
    private string statusText = "Загрузка данных...";

    [ObservableProperty]
    private string reportText = "";

    private async Task LoadAsync()
    {
        try
        {
            var items = await _itemService.LoadItemsAsync();
            var categories = await _categoryService.LoadCategoriesAsync();
            var users = await _adminService.LoadUsersAsync();

            // Запоминаем выделенные элементы: коллекции пересоздаются.
            var selectedItemId = SelectedItem?.Id;
            var selectedUserId = SelectedUser?.Id;

            Items.Clear();

            foreach (var item in items)
            {
                Items.Add(item);
            }

            Categories.Clear();

            foreach (var category in categories)
            {
                Categories.Add(category);
            }

            Users.Clear();

            foreach (var user in users)
            {
                Users.Add(user);
            }

            if (selectedItemId is not null)
            {
                SelectedItem = Items.FirstOrDefault(item => item.Id == selectedItemId.Value);
            }

            if (selectedUserId is not null)
            {
                SelectedUser = Users.FirstOrDefault(user => user.Id == selectedUserId.Value);
            }

            if (!Roles.Contains("Гость"))
            {
                Roles.Add("Гость");
            }

            if (!Roles.Contains("Покупатель"))
            {
                Roles.Add("Покупатель");
            }

            if (!Roles.Contains("Менеджер"))
            {
                Roles.Add("Менеджер");
            }

            if (!Roles.Contains("Администратор"))
            {
                Roles.Add("Администратор");
            }

            StatusText = $"Товаров: {Items.Count}, пользователей: {Users.Count}";
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(ex);

            StatusText = "Не удалось загрузить данные. Проверьте подключение к базе данных.";
        }
    }

    partial void OnSelectedItemChanged(Product? value)
    {
        if (value is null)
        {
            return;
        }

        ItemName = value.Name_Item;
        ItemArticule = value.NameArticule;
        ItemPrice = value.Price.ToString("0.00");
        ItemQuantity = value.Quantity.ToString();

        SelectedItemCategory =
            Categories.FirstOrDefault(category => category.Id == value.Category_Id);
    }

    [RelayCommand]
    private async Task AddItemAsync()
    {
        var categoryId = SelectedItemCategory?.Id ?? 0;

        try
        {
            await _adminService.AddItemAsync(
                ItemName,
                ItemArticule,
                categoryId,
                ParsePrice(),
                ParseQuantity());

            await ReloadAsync();

            StatusText = $"Товар «{ItemName}» добавлен в каталог.";
        }
        catch (ArgumentException ex)
        {
            StatusText = ex.Message;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(ex);

            StatusText = "Не удалось добавить товар.";
        }
    }

    [RelayCommand]
    private async Task UpdateItemAsync()
    {
        if (SelectedItem is null)
        {
            StatusText = "Выберите товар для изменения.";

            return;
        }

        var categoryId = SelectedItemCategory?.Id ?? SelectedItem.Category_Id;

        try
        {
            await _adminService.UpdateItemAsync(
                SelectedItem.Id,
                ItemName,
                ItemArticule,
                categoryId,
                ParsePrice(),
                ParseQuantity());

            await ReloadAsync();

            StatusText = $"Товар «{ItemName}» обновлён.";
        }
        catch (ArgumentException ex)
        {
            StatusText = ex.Message;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(ex);

            StatusText = "Не удалось изменить товар.";
        }
    }

    [RelayCommand]
    private async Task DeleteItemAsync(Product? item)
    {
        var target = item ?? SelectedItem;

        if (target is null)
        {
            StatusText = "Выберите товар для удаления.";

            return;
        }

        try
        {
            await _adminService.DeleteItemAsync(target.Id);

            await ReloadAsync();

            StatusText = $"Товар «{target.Name_Item}» удалён из каталога.";
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(ex);

            StatusText = ex.Message;
        }
    }

    [RelayCommand]
    private async Task AddCategoryAsync()
    {
        try
        {
            await _adminService.AddCategoryAsync(NewCategoryName);

            NewCategoryName = "";

            await ReloadAsync();

            StatusText = "Категория добавлена.";
        }
        catch (ArgumentException ex)
        {
            StatusText = ex.Message;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(ex);

            StatusText = "Не удалось добавить категорию.";
        }
    }

    [RelayCommand]
    private async Task DeleteCategoryAsync(CategoryOption? category)
    {
        if (category is null)
        {
            StatusText = "Выберите категорию для удаления.";

            return;
        }

        try
        {
            await _adminService.DeleteCategoryAsync(category.Id);

            await ReloadAsync();

            StatusText = $"Категория «{category.Name}» удалена.";
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(ex);

            StatusText = ex.Message;
        }
    }

    [RelayCommand]
    private async Task ChangeRoleAsync()
    {
        if (SelectedUser is null)
        {
            StatusText = "Выберите пользователя.";

            return;
        }

        // Значения сохраняем до перезагрузки списка.
        var userId = SelectedUser.Id;
        var login = SelectedUser.Login;
        var role = SelectedRole;

        try
        {
            await _adminService.ChangeUserRoleAsync(userId, role);

            await ReloadAsync();

            StatusText = $"Пользователю {login} назначена роль «{role}».";
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(ex);

            StatusText = ex.Message;
        }
    }

    [RelayCommand]
    private async Task BuildReportAsync()
    {
        try
        {
            var report = await _adminService.BuildReportAsync();

            ReportLines.Clear();

            foreach (var line in report.ByStatus)
            {
                ReportLines.Add(line);
            }

            ReportText =
                $"Заказов: {report.OrdersCount}; " +
                $"товаров: {report.ItemsCount}; " +
                $"пользователей: {report.UsersCount}; " +
                $"отзывов: {report.ReviewsCount}; " +
                $"отменено: {report.CancelledCount}; " +
                $"сумма заказов: {report.TotalAmount:N2} руб.";
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(ex);

            ReportText = "Не удалось построить отчёт.";
        }
    }

    private async Task ReloadAsync()
    {
        await LoadAsync();

        await _mainViewModel.ReloadCatalogAsync();
    }

    private decimal ParsePrice()
    {
        return decimal.TryParse(
            ItemPrice.Replace(',', '.'),
            System.Globalization.NumberStyles.Any,
            System.Globalization.CultureInfo.InvariantCulture,
            out var value)
            ? value
            : 0m;
    }

    private int ParseQuantity()
    {
        return int.TryParse(ItemQuantity, out var value) ? value : 0;
    }
}
