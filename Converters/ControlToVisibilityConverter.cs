using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;
using WordBarcodeStudio.Models;

namespace WordBarcodeStudio.Converters;

[ValueConversion(typeof(BarcodeControlType), typeof(Visibility))]
public class ControlToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is BarcodeControlType ct && parameter is string p)
        {
            return ct.ToString() == p ? Visibility.Visible : Visibility.Collapsed;
        }
        return Visibility.Collapsed;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotSupportedException();
}
