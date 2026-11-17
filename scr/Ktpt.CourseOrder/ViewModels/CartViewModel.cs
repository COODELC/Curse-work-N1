using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Ktpt.CourseOrder.Models;
using System.Collections.ObjectModel;
using System.Threading.Tasks;

namespace Ktpt.CourseOrder.ViewModels;

/// <summary>
/// Окно корзины.
/// </summary>
public partial class CartViewModel : ObservableObject
{
    private readonly MainViewModel _mainViewModel;

    public CartViewModel(MainViewModel mainViewModel)
    {
        _mainViewModel = mainViewModel;

        Cart = mainViewModel.Cart;

        mainViewModel.CartChanged += OnCartChanged;
    }

    public ObservableCollection<CartItem> Cart { get; }

    public decimal CartTotal => _mainViewModel.CartTotal;

    [ObservableProperty]
    private string paymentStatus = "";

    private void OnCartChanged() => OnPropertyChanged(nameof(CartTotal));

    [RelayCommand]
    private void Increase(CartItem? item) => _mainViewModel.IncreaseQuantityCommand.Execute(item);

    [RelayCommand]
    private void Decrease(CartItem? item) => _mainViewModel.DecreaseQuantityCommand.Execute(item);

    [RelayCommand]
    private void RemoveFromCart(CartItem? item) => _mainViewModel.RemoveFromCartCommand.Execute(item);

    [RelayCommand]
    private async Task PayAsync()
    {
        var (_, message) = await _mainViewModel.TryPayAsync();

        PaymentStatus = message;
    }
}
