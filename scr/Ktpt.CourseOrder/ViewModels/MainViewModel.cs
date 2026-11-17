using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Ktpt.CourseOrder.Auth;
using Ktpt.CourseOrder.Models;
using Ktpt.CourseOrder.Services;
using Ktpt.CourseOrder.Views;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;

namespace Ktpt.CourseOrder.ViewModels;

/// <summary>
/// Главная форма: каталог товаров, поиск, фильтрация по категориям, корзина.
/// </summary>
public partial class MainViewModel : ObservableObject
{
    private readonly ItemService _itemService;

    private readonly CategoryService _categoryService;

    private readonly OrderService _orderService;

    private readonly ReviewService _reviewService;

    private CartWindow? _cartWindow;

    private OrdersWindow? _ordersWindow;

    private ManagerWindow? _managerWindow;

    private AdminWindow? _adminWindow;

    public ObservableCollection<Product> Products { get; } = new();

    public ObservableCollection<CategoryOption> Categories { get; } = new();

    public ObservableCollection<CartItem> Cart { get; } = new();

    [ObservableProperty]
    private string statusText = "Загрузка каталога...";

    [ObservableProperty]
    private string searchText = "";

    [ObservableProperty]
    private CategoryOption? selectedCategory;

    /// <summary>Главное окно — владелец диалоговых окон.</summary>
    public Avalonia.Controls.Window? Owner { get; set; }

    /// <summary>Событие изменения корзины: окно корзины пересчитывает итог.</summary>
    public event Action? CartChanged;

    public decimal CartTotal => Cart.Sum(item => item.Total);

    public MainViewModel()
    {
        _itemService = new ItemService(AppServices.Data);
        _categoryService = new CategoryService(AppServices.Data);
        _orderService = new OrderService(AppServices.Data);
        _reviewService = new ReviewService(AppServices.Data);

        // Окна, открытые до входа, закрываются, а каталог обновляется.
        AppServices.Session.Changed += OnSessionChanged;

        _ = InitializeAsync();
    }

    /// <summary>
    /// Реакция на вход и выход пользователя: закрываются вспомогательные окна,
    /// очищается корзина, обновляется каталог.
    /// </summary>
    private void OnSessionChanged()
    {
        _cartWindow?.Close();
        _ordersWindow?.Close();
        _managerWindow?.Close();
        _adminWindow?.Close();

        ClearCart();

        var user = AppServices.Session.CurrentUser;

        StatusText = user is null
            ? "Вы вышли из аккаунта. Корзина очищена."
            : $"Вы вошли как {user.FirstName}.";

        _ = ReloadAsync();
    }

    private async Task InitializeAsync()
    {
        try
        {
            await LoadCategoriesAsync();
            await LoadProductsAsync();
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(ex);

            StatusText =
                "Не удалось загрузить каталог. " +
                "Проверьте подключение к базе данных.";
        }
    }

    private async Task LoadCategoriesAsync()
    {
        var categories = await _categoryService.LoadCategoriesAsync();

        Categories.Clear();

        foreach (var category in categories)
        {
            Categories.Add(category);
        }
    }

    private async Task LoadProductsAsync()
    {
        var products = await _itemService.LoadItemsAsync(
            SelectedCategory?.Id,
            SearchText);

        Products.Clear();

        foreach (var product in products)
        {
            Products.Add(product);
        }

        UpdateStatus(products.Count);
    }

    private void UpdateStatus(int count)
    {
        var query = (SearchText ?? "").Trim();

        StatusText = string.IsNullOrWhiteSpace(query)
            ? $"Всего товаров: {count}"
            : count == 0
                ? $"По запросу «{query}» ничего не найдено."
                : $"Найдено товаров: {count}";
    }

    [RelayCommand]
    private async Task SearchAsync()
    {
        await ReloadAsync();
    }

    [RelayCommand]
    private async Task ResetSearchAsync()
    {
        SearchText = "";

        SelectedCategory = null;

        await ReloadAsync();
    }

    public async Task ReloadAsync()
    {
        try
        {
            await LoadProductsAsync();
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(ex);

            StatusText = "Не удалось загрузить каталог. Проверьте подключение к базе данных.";
        }
    }

    [RelayCommand]
    private void AddToCart(Product? product)
    {
        if (!AppServices.Session.Can(Permission.PlaceOrder))
        {
            StatusText = "Для добавления товара в корзину войдите в аккаунт.";

            return;
        }

        if (product is null)
        {
            return;
        }

        var catalogProduct = Products.FirstOrDefault(item => item.Id == product.Id)
            ?? product;

        if (catalogProduct.Quantity <= 0)
        {
            StatusText = "Товар недоступен для заказа.";

            return;
        }

        var existing = Cart.FirstOrDefault(item => item.ProductId == catalogProduct.Id);

        var quantityInCart = existing?.Quantity ?? 0;

        if (quantityInCart >= catalogProduct.Quantity)
        {
            StatusText =
                $"Нельзя добавить больше товара «{catalogProduct.Name_Item}». " +
                $"В наличии: {catalogProduct.Quantity} шт.";

            return;
        }

        if (existing is null)
        {
            Cart.Add(new CartItem
            {
                ProductId = catalogProduct.Id,
                Name_Item = catalogProduct.Name_Item,
                Price = catalogProduct.Price,
                Quantity = 1
            });
        }
        else
        {
            existing.Quantity += 1;
        }

        NotifyCartChanged();

        StatusText = $"Товар «{catalogProduct.Name_Item}» добавлен в корзину.";
    }

    [RelayCommand]
    private void IncreaseQuantity(CartItem? item)
    {
        if (item is null)
        {
            return;
        }

        var product = Products.FirstOrDefault(item2 => item2.Id == item.ProductId);

        var available = product?.Quantity ?? item.Quantity;

        if (item.Quantity >= available)
        {
            StatusText =
                $"Нельзя добавить больше товара «{item.Name_Item}». " +
                $"В наличии: {available} шт.";

            return;
        }

        item.Quantity += 1;

        NotifyCartChanged();
    }

    [RelayCommand]
    private void DecreaseQuantity(CartItem? item)
    {
        if (item is null)
        {
            return;
        }

        if (item.Quantity <= 1)
        {
            RemoveFromCart(item);

            return;
        }

        item.Quantity -= 1;

        NotifyCartChanged();
    }

    [RelayCommand]
    private void RemoveFromCart(CartItem? item)
    {
        if (item is null)
        {
            return;
        }

        Cart.Remove(item);

        NotifyCartChanged();

        StatusText = "Товар удалён из корзины.";
    }

    private void NotifyCartChanged()
    {
        OnPropertyChanged(nameof(CartTotal));

        CartChanged?.Invoke();
    }

    public void ClearCart()
    {
        Cart.Clear();

        NotifyCartChanged();

        StatusText = "Корзина очищена.";
    }

    /// <summary>
    /// Оформление заказа: записи сохраняются в таблицу Orders,
    /// остатки списываются в той же транзакции.
    /// </summary>
    public async Task<(bool Success, string Message)> TryPayAsync()
    {
        if (!AppServices.Session.Can(Permission.PlaceOrder))
        {
            return (false, "Для оформления заказа войдите в аккаунт.");
        }

        if (Cart.Count == 0)
        {
            return (false, "Корзина пуста.");
        }

        if (Cart.Any(item => item.Quantity <= 0))
        {
            return (false, "Количество товара должно быть больше нуля.");
        }

        var user = AppServices.Session.CurrentUser;

        if (user is null)
        {
            return (false, "Для оформления заказа войдите в аккаунт.");
        }

        var lines = Cart
            .Select(item => (ProductId: item.ProductId, Quantity: item.Quantity))
            .ToList();

        OrderResult result;

        try
        {
            result = await _orderService.CreateOrderAsync(user.Id, lines);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(ex);

            var message = "Не удалось оформить заказ. Проверьте подключение к базе данных.";

            StatusText = message;

            return (false, message);
        }

        if (!result.Success)
        {
            StatusText = result.Message;

            return (false, result.Message);
        }

        await ReloadAsync();

        ClearCart();

        var text = result.IsLarge
            ? $"{result.Message} Заказ признан крупным (более {OrderService.LargeOrderThreshold} позиций)."
            : result.Message;

        if (result.DiscountTotal > 0)
        {
            text += $" Применена скидка {OrderService.DiscountPercent:N0}%, " +
                    $"выгода {result.DiscountTotal:N2} руб.";
        }

        text += " Банковская оплата не выполнялась.";

        StatusText = text;

        return (true, text);
    }

    /// <summary>
    /// Добавление отзыва о товаре.
    /// </summary>
    public async Task<string> AddReviewAsync(int itemId, string text)
    {
        if (!AppServices.Session.Can(Permission.AddReview))
        {
            return "Оставлять отзывы может только авторизованный пользователь.";
        }

        var user = AppServices.Session.CurrentUser;

        try
        {
            await _reviewService.AddReviewAsync(itemId, text, user?.Id);

            await ReloadAsync();

            StatusText = user is null
                ? "Отзыв сохранён."
                : $"Отзыв сохранён. Автор: {user.FirstName}.";

            return StatusText;
        }
        catch (ArgumentException ex)
        {
            return ex.Message;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(ex);

            return "Не удалось сохранить отзыв. Проверьте подключение к базе данных.";
        }
    }

    /// <summary>
    /// Показывает список отзывов о товаре вместе с их авторами.
    /// </summary>
    [RelayCommand]
    private async Task OpenReviewsAsync(Product? product)
    {
        if (product is null)
        {
            return;
        }

        try
        {
            var reviews = await _reviewService.GetReviewsAsync(product.Id);

            var window = new ReviewsWindow(product.Name_Item, reviews);

            if (Owner is null)
            {
                window.Show();

                return;
            }

            await window.ShowDialog(Owner);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(ex);

            StatusText = "Не удалось загрузить отзывы. Проверьте подключение к базе данных.";
        }
    }

    [RelayCommand]
    private async Task OpenReviewAsync(Product? product)
    {
        if (product is null)
        {
            return;
        }

        if (!AppServices.Session.Can(Permission.AddReview))
        {
            StatusText = "Оставлять отзывы может только авторизованный пользователь.";

            return;
        }

        var user = AppServices.Session.CurrentUser;

        var window = new ReviewWindow(product.Name_Item, user?.FirstName);

        var text = Owner is null
            ? null
            : await window.ShowDialog<string>(Owner);

        if (string.IsNullOrWhiteSpace(text))
        {
            return;
        }

        var message = await AddReviewAsync(product.Id, text);

        StatusText = message;
    }

    /// <summary>
    /// Открывает окно с подробным описанием товара.
    /// </summary>
    [RelayCommand]
    private async Task OpenProductDetailsAsync(Product? product)
    {
        if (product is null)
        {
            return;
        }

        var window = new ProductDetailsWindow(product);

        if (Owner is null)
        {
            window.Show();

            return;
        }

        await window.ShowDialog(Owner);
    }

    [RelayCommand]
    private void OpenCart()
    {
        if (!AppServices.Session.Can(Permission.PlaceOrder))
        {
            StatusText = "Для работы с корзиной войдите в аккаунт.";

            return;
        }

        if (_cartWindow is not null)
        {
            _cartWindow.Activate();

            return;
        }

        var window = new CartWindow(this);

        _cartWindow = window;

        window.Closed += (_, _) => _cartWindow = null;

        window.Show();
    }

    [RelayCommand]
    private void OpenMyOrders()
    {
        if (!AppServices.Session.Can(Permission.ViewOwnOrders))
        {
            StatusText = "История покупок доступна только авторизованному пользователю.";

            return;
        }

        if (_ordersWindow is not null)
        {
            _ordersWindow.Activate();

            return;
        }

        var window = new OrdersWindow(this);

        _ordersWindow = window;

        window.Closed += (_, _) => _ordersWindow = null;

        window.Show();
    }

    [RelayCommand]
    private void OpenManagerOrders()
    {
        if (!AppServices.Session.Can(Permission.ManageOrders))
        {
            StatusText = "Обработка заказов доступна сотруднику и администратору.";

            return;
        }

        if (_managerWindow is not null)
        {
            _managerWindow.Activate();

            return;
        }

        var window = new ManagerWindow(this);

        _managerWindow = window;

        window.Closed += (_, _) => _managerWindow = null;

        window.Show();
    }

    [RelayCommand]
    private void OpenAdmin()
    {
        if (!AppServices.Session.Can(Permission.ManageProducts))
        {
            StatusText = "Управление каталогом доступно администратору.";

            return;
        }

        if (_adminWindow is not null)
        {
            _adminWindow.Activate();

            return;
        }

        var window = new AdminWindow(this);

        _adminWindow = window;

        window.Closed += (_, _) => _adminWindow = null;

        window.Show();
    }

    public void HandleSignOut()
    {
        OnSessionChanged();
    }

    /// <summary>
    /// Повторная загрузка каталога после изменений в окне администратора.
    /// </summary>
    public async Task ReloadCatalogAsync()
    {
        await ReloadAsync();
    }
}
