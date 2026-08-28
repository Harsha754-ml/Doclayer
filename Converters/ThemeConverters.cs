using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;
using MediaColor = System.Windows.Media.Color;
using MediaBrush = System.Windows.Media.Brush;
using SolidColorBrush = System.Windows.Media.SolidColorBrush;

namespace WordBarcodeStudio.Converters;

[ValueConversion(typeof(string), typeof(Visibility))]
public class ViewEqualsVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is string currentView && parameter is string targetView)
        {
            return currentView.Equals(targetView, StringComparison.OrdinalIgnoreCase) 
                ? Visibility.Visible 
                : Visibility.Collapsed;
        }
        return Visibility.Collapsed;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotSupportedException();
}

[ValueConversion(typeof(string), typeof(bool))]
public class ViewEqualsBoolConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is string currentView && parameter is string targetView)
        {
            return currentView.Equals(targetView, StringComparison.OrdinalIgnoreCase);
        }
        return false;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotSupportedException();
}

[ValueConversion(typeof(string), typeof(MediaBrush))]
public class StatusToBrushConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is string status)
        {
            var s = status.ToLowerInvariant();
            if (s.Contains("connect") && !s.Contains("disconnect") || s.Contains("complete") || s.Contains("success"))
            {
                return new SolidColorBrush(MediaColor.FromRgb(0x22, 0xC5, 0x5E)); // Green
            }
            if (s.Contains("start") || s.Contains("creat") || s.Contains("insert") || s.Contains("updat") || s.Contains("generat"))
            {
                return new SolidColorBrush(MediaColor.FromRgb(0xF5, 0x9E, 0x0B)); // Amber/Orange
            }
            if (s.Contains("fail") || s.Contains("error"))
            {
                return new SolidColorBrush(MediaColor.FromRgb(0xEF, 0x44, 0x44)); // Red
            }
            if (s.Contains("ready"))
            {
                return new SolidColorBrush(MediaColor.FromRgb(0x3B, 0x82, 0xF6)); // Blue
            }
        }
        return new SolidColorBrush(MediaColor.FromRgb(0x64, 0x74, 0x8B)); // Slate/Muted
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotSupportedException();
}

[ValueConversion(typeof(string), typeof(Visibility))]
public class StringNotEmptyToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        bool notEmpty = !string.IsNullOrWhiteSpace(value as string);
        if (parameter is string s && s.Equals("invert", StringComparison.OrdinalIgnoreCase))
            notEmpty = !notEmpty;
        return notEmpty ? Visibility.Visible : Visibility.Collapsed;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotSupportedException();
}
