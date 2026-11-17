using Avalonia.Controls;
using Avalonia.Interactivity;

namespace Ktpt.CourseOrder.Views;

/// <summary>
/// Диалог ввода отзыва о товаре.
/// </summary>
public partial class ReviewWindow : Window
{
    public ReviewWindow(string itemName, string? authorName = null)
    {
        InitializeComponent();

        CaptionText.Text = $"Отзыв о товаре «{itemName}»";

        AuthorText.Text = authorName is null
            ? "Отзыв будет сохранён от имени гостя."
            : $"Отзыв будет сохранён от имени «{authorName}».";

        AuthorText.IsVisible = true;
    }

    private void SaveButton_Click(object? sender, RoutedEventArgs e)
    {
        Close(ReviewBox.Text ?? string.Empty);
    }

    private void CancelButton_Click(object? sender, RoutedEventArgs e)
    {
        Close(null);
    }
}
