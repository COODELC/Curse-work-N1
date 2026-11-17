using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Ktpt.CourseOrder.Models;
using Ktpt.CourseOrder.Services;
using System;
using System.Collections.ObjectModel;
using System.Threading.Tasks;

namespace Ktpt.CourseOrder.ViewModels;

/// <summary>
/// История покупок пользователя.
/// </summary>
public partial class OrdersViewModel : ObservableObject
{
    private readonly MainViewModel _mainViewModel;

    private readonly OrderService _orderService;

    public OrdersViewModel(MainViewModel mainViewModel)
    {
        _mainViewModel = mainViewModel;

        _orderService = new OrderService(AppServices.Data);

        _ = LoadAsync();
    }

    public ObservableCollection<OrderInfo> Orders { get; } = new();

    [ObservableProperty]
    private string statusText = "Загрузка заказов...";

    private async Task LoadAsync()
    {
        var user = AppServices.Session.CurrentUser;

        if (user is null)
        {
            StatusText = "Войдите в аккаунт, чтобы увидеть историю покупок.";

            return;
        }

        try
        {
            var orders = await _orderService.GetUserOrdersAsync(user.Id);

            Orders.Clear();

            foreach (var order in orders)
            {
                Orders.Add(order);
            }

            StatusText = $"Всего заказов: {Orders.Count}";
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(ex);

            StatusText = "Не удалось загрузить заказы. Проверьте подключение к базе данных.";
        }
    }

    [RelayCommand]
    private async Task CancelAsync(OrderInfo? order)
    {
        if (order is null)
        {
            return;
        }

        try
        {
            var cancelled = await _orderService.CancelOrderAsync(order.Id);

            await LoadAsync();

            await _mainViewModel.ReloadCatalogAsync();

            StatusText = cancelled
                ? $"Заказ № {order.Id} отменён, товар возвращён на склад."
                : $"Заказ № {order.Id} уже был отменён ранее.";
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(ex);

            StatusText = "Не удалось отменить заказ.";
        }
    }
}
