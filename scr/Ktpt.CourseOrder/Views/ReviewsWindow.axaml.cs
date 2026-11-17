using Avalonia.Controls;
using Avalonia.Interactivity;
using Ktpt.CourseOrder.Models;
using System.Collections.Generic;

namespace Ktpt.CourseOrder.Views;

/// <summary>
/// Окно со списком отзывов о товаре: текст отзыва и его автор.
/// </summary>
public partial class ReviewsWindow : Window
{
    public ReviewsWindow(string itemName, List<ProductReview> reviews)
    {
        InitializeComponent();

        CaptionText.Text = $"Отзывы о товаре «{itemName}»";

        CountText.Text = reviews.Count > 0
            ? $"Всего отзывов: {reviews.Count}"
            : "Отзывов пока нет";

        ReviewsList.ItemsSource = reviews;

        EmptyText.Text = "Об этом товаре ещё никто не оставил отзыв.\n"
                       + "Станьте первым — ваш отзыв увидят все покупатели.";

        EmptyText.IsVisible = reviews.Count == 0;
    }

    private void CloseButton_Click(object? sender, RoutedEventArgs e)
    {
        Close();
    }
}
