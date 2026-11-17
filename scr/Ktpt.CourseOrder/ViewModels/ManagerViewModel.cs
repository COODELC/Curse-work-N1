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
/// Обработка заказов сотрудником магазина.
/// </summary>
public partial class ManagerViewModel : ObservableObject
{
    private readonly MainViewModel _mainViewModel;

    private readonly OrderService _orderService;

    public ManagerViewModel(MainViewModel mainViewModel)
    {
        _mainViewModel = mainViewModel;

        _orderService = new OrderService(AppServices.Data);

        _ = LoadAsync();
    }

    public ObservableCollection<OrderInfo> Orders { get; } = new();

    public ObservableCollection<OrderStatusOption> Statuses { get; } = new();

    [ObservableProperty]
    private OrderInfo? selectedOrder;

    [ObservableProperty]
    private OrderStatusOption? selectedStatus;

    [ObservableProperty]
    private string statusText = "Загрузка заказов...";

    private async Task LoadAsync()
    {
        try
        {
            var orders = await _orderService.GetAllOrdersAsync();

            var statuses = await _orderService.LoadStatusesAsync();

            // Запоминаем выделенные элементы: коллекции пересоздаются.
            var selectedOrderId = SelectedOrder?.Id;
            var selectedStatusId = SelectedStatus?.Id;

            Orders.Clear();

            foreach (var order in orders)
            {
                Orders.Add(order);
            }

            Statuses.Clear();

            foreach (var status in statuses)
            {
                Statuses.Add(status);
            }

            if (selectedOrderId is not null)
            {
                SelectedOrder = Orders.FirstOrDefault(order => order.Id == selectedOrderId.Value);
            }

            if (selectedStatusId is not null)
            {
                SelectedStatus = Statuses.FirstOrDefault(status => status.Id == selectedStatusId.Value);
            }

            StatusText = $"Заказов в системе: {Orders.Count}";
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(ex);

            StatusText = "Не удалось загрузить заказы. Проверьте подключение к базе данных.";
        }
    }

    [RelayCommand]
    private async Task ChangeStatusAsync()
    {
        if (SelectedOrder is null || SelectedStatus is null)
        {
            StatusText = "Выберите заказ и новый статус.";

            return;
        }

        // Значения сохраняем до перезагрузки списка.
        var orderId = SelectedOrder.Id;
        var statusId = SelectedStatus.Id;
        var statusName = SelectedStatus.Name;

        try
        {
            await _orderService.SetStatusAsync(orderId, statusId);

            await LoadAsync();

            await _mainViewModel.ReloadCatalogAsync();

            StatusText = $"Заказу № {orderId} присвоен статус «{statusName}».";
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(ex);

            StatusText = "Не удалось изменить статус заказа.";
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

    /// <summary>
    /// Количество заказов в выбранном статусе — для сводки.
    /// </summary>
    public int CountByStatus(int statusId) =>
        Orders.Count(order => order.StatusId == statusId);
}
