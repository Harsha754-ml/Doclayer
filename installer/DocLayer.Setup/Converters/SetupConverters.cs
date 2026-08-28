using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media;

namespace DocLayer.Setup.Converters;

public class BoolToVisConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        bool b = value is bool v && v;
        bool invert = parameter is string s && s.Equals("invert", StringComparison.OrdinalIgnoreCase);
        if (invert) b = !b;
        return b ? Visibility.Visible : Visibility.Collapsed;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotImplementedException();
}

public class StepHighlightConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        bool isActive = value is bool b && b;
        return isActive
            ? new SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 255, 255))
            : new SolidColorBrush(System.Windows.Media.Color.FromRgb(161, 161, 170));
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotImplementedException();
}

public class StepBadgeBgConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        bool isActive = value is bool b && b;
        return isActive
            ? new SolidColorBrush(System.Windows.Media.Color.FromRgb(244, 244, 245))
            : new SolidColorBrush(System.Windows.Media.Color.FromRgb(30, 30, 30));
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotImplementedException();
}

public class StepBadgeFgConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        bool isActive = value is bool b && b;
        return isActive
            ? new SolidColorBrush(System.Windows.Media.Color.FromRgb(9, 9, 11))
            : new SolidColorBrush(System.Windows.Media.Color.FromRgb(161, 161, 170));
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotImplementedException();
}
