using Avalonia.Data.Converters;
using Avalonia.Media;
using System;
using System.Globalization;

namespace Ktpt.CourseOrder.Converters;

/// <summary>
/// Цвет плашки со статусом заказа.
/// Готовые заказы — зелёные, отменённые — красные, остальные — оттенки синего.
/// </summary>
public sealed class StatusBrushConverter : IValueConverter
{
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        var status = value as string ?? string.Empty;

        var color = status switch
        {
            "Доставлен" => "#059669",
            "Передан в доставку" => "#0891B2",
            "Собирается" => "#D97706",
            "Подтверждён" => "#4F46E5",
            "Оформление заказа" => "#7C3AED",
            "Отменён" => "#E11D48",
            _ => "#64748B"
        };

        return new SolidColorBrush(Color.Parse(color));
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        throw new NotSupportedException();
    }
}

/// <summary>
/// Цвет плашки категории товара: у каждой категории свой оттенок.
/// </summary>
public sealed class CategoryBrushConverter : IValueConverter
{
    /// <summary>Закреплённые цвета категорий каталога.</summary>
    private static readonly Dictionary<string, string> KnownCategories = new()
    {
        ["Двигатель"] = "#4F46E5",
        ["Тормозная система"] = "#DB2777",
        ["Подвеска"] = "#0891B2",
        ["Электрика"] = "#059669",
        ["Масла и жидкости"] = "#D97706",
        ["Кузовные детали"] = "#7C3AED"
    };

    private static readonly string[] Palette =
    {
        "#4F46E5",
        "#DB2777",
        "#0891B2",
        "#059669",
        "#D97706",
        "#7C3AED"
    };

    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        var category = value as string ?? string.Empty;

        if (category.Length == 0)
        {
            return new SolidColorBrush(Color.Parse("#64748B"));
        }

        if (KnownCategories.TryGetValue(category, out var known))
        {
            return new SolidColorBrush(Color.Parse(known));
        }

        // Для неизвестной категории цвет выбирается устойчивым хешем:
        // беззнаковые вычисления исключают переполнение.
        uint hash = 2166136261;

        foreach (var symbol in category)
        {
            hash = unchecked((hash ^ symbol) * 16777619);
        }

        return new SolidColorBrush(Color.Parse(Palette[hash % (uint)Palette.Length]));
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        throw new NotSupportedException();
    }
}
