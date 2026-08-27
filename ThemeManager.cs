using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Media;

namespace WordBarcodeStudio;

/// <summary>
/// Swaps the application's colour resources between a light and a dark palette.
/// The XAML binds to these resources via DynamicResource, so changing them here
/// re-themes the whole UI at runtime.
/// </summary>
internal static class ThemeManager
{
    public static bool IsDark { get; private set; }

    private static readonly Dictionary<string, (string Light, string Dark)> Palette = new()
    {
        ["AppBg"] = ("#F4F6FB", "#0F172A"),
        ["CardBg"] = ("#FFFFFF", "#1E293B"),
        ["CardBorder"] = ("#E3E8F0", "#334155"),
        ["FieldLabelFg"] = ("#475569", "#CBD5E1"),
        ["PrimaryBtnBg"] = ("#2563EB", "#3B82F6"),
        ["PrimaryBtnFg"] = ("#FFFFFF", "#FFFFFF"),
        ["SecondaryBtnBg"] = ("#EEF2FF", "#334155"),
        ["SecondaryBtnFg"] = ("#1E3A8A", "#E2E8F0"),
        ["TextBoxBg"] = ("#FFFFFF", "#0F172A"),
        ["ControlBorder"] = ("#CBD5E1", "#475569"),
        ["HeaderCardBg"] = ("#2563EB", "#1E3A8A"),
        ["HeaderText"] = ("#FFFFFF", "#FFFFFF"),
        ["HeaderSub"] = ("#DBEAFE", "#BFDBFE"),
        ["PreviewBg"] = ("#F8FAFC", "#0B1220"),
        ["MutedText"] = ("#94A3B8", "#64748B"),
        ["StatusBg"] = ("#0F172A", "#020617"),
        ["StatusFg"] = ("#E2E8F0", "#E2E8F0"),
        ["InputText"] = ("#1E293B", "#E2E8F0"),
    };

    public static void Initialize() => Apply(LoadPreference());

    public static void Toggle() => Apply(!IsDark);

    public static void Apply(bool dark)
    {
        IsDark = dark;
        var resources = System.Windows.Application.Current.Resources;
        foreach (var kvp in Palette)
        {
            var color = dark ? kvp.Value.Dark : kvp.Value.Light;
            resources[kvp.Key] = new SolidColorBrush(
                (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(color));
        }

        if (System.Windows.Application.Current.MainWindow is Window window)
        {
            window.Background = (SolidColorBrush)resources["AppBg"];
        }

        SavePreference(dark);
    }

    private static bool LoadPreference()
    {
        try
        {
            var path = PreferenceFile();
            if (File.Exists(path))
            {
                return File.ReadAllText(path).Trim().Equals("Dark", StringComparison.OrdinalIgnoreCase);
            }
        }
        catch
        {
        }

        return false;
    }

    private static void SavePreference(bool dark)
    {
        try
        {
            var dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Doclayer");
            Directory.CreateDirectory(dir);
            File.WriteAllText(Path.Combine(dir, "theme.txt"), dark ? "Dark" : "Light");
        }
        catch
        {
        }
    }

    private static string PreferenceFile()
    {
        return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Doclayer", "theme.txt");
    }
}
