using Avalonia.Data.Converters;
using Avalonia.Media.Imaging;
using Avalonia.Platform;
using System;
using System.Globalization;

namespace Ktpt.CourseOrder.Converters;

/// <summary>
/// Преобразует путь к изображению товара в готовую картинку.
/// В базе путь хранится как "/Assets/Products/battery.png", поэтому
/// к нему добавляется префикс avares:// с именем сборки.
/// Если файла нет, возвращается null — в карточке останется фоновая заливка.
/// </summary>
public sealed class ImagePathConverter : IValueConverter
{
    /// <summary>Имя сборки, в которой лежат ресурсы Assets.</summary>
    private const string AssemblyName = "Ktpt.CourseOrder";

    public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is Bitmap bitmap)
        {
            return bitmap;
        }

        return LoadBitmap(value as string);
    }

    public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        throw new NotSupportedException();
    }

    private static Bitmap? LoadBitmap(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return null;
        }

        var uri = BuildUri(path.Trim());

        if (uri is null || !AssetLoader.Exists(uri))
        {
            return null;
        }

        try
        {
            using var stream = AssetLoader.Open(uri);

            return new Bitmap(stream);
        }
        catch (Exception)
        {
            // Битая картинка не должна останавливать загрузку каталога.
            return null;
        }
    }

    private static Uri? BuildUri(string path)
    {
        // Путь уже задан как avares:// — используем его как есть.
        if (path.StartsWith("avares://", StringComparison.OrdinalIgnoreCase))
        {
            return new Uri(path);
        }

        var relative = path.TrimStart('/');

        return new Uri($"avares://{AssemblyName}/{relative}");
    }
}
