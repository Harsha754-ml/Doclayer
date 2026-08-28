using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Media;

namespace WordBarcodeStudio;

/// <summary>
/// Swaps the application's colour resources between a sleek pitch-black dark mode
/// and a clean modern light mode.
/// </summary>
internal static class ThemeManager
{
    public static bool IsDark { get; private set; }

    private static readonly Dictionary<string, (string Light, string Dark)> Palette = new()
    {
        // Application surfaces (Sleek neutral dark / pitch black in Dark mode)
        ["AppBg"] = ("#F8FAFC", "#0A0A0A"),
        ["SidebarBg"] = ("#FFFFFF", "#121212"),
        ["SidebarBorder"] = ("#E4E4E7", "#1E1E1E"),
        ["HeaderBg"] = ("#FFFFFF", "#121212"),
        ["HeaderBorder"] = ("#E4E4E7", "#1E1E1E"),
        ["CardBg"] = ("#FFFFFF", "#161616"),
        ["CardBorder"] = ("#E4E4E7", "#242424"),
        ["CardBgElevated"] = ("#F4F4F5", "#1C1C1C"),
        ["CardBorderSubtle"] = ("#E4E4E7", "#202020"),

        // Typography & Text
        ["TextPrimary"] = ("#09090B", "#FFFFFF"),
        ["TextSecondary"] = ("#52525B", "#A1A1AA"),
        ["MutedText"] = ("#71717A", "#71717A"),
        ["FieldLabelFg"] = ("#27272A", "#A1A1AA"),
        ["InputText"] = ("#09090B", "#FFFFFF"),

        // Brand & Neutral Accents
        ["PrimaryBrand"] = ("#18181B", "#FFFFFF"),
        ["AccentPurple"] = ("#71717A", "#A1A1AA"),
        ["AccentCyan"] = ("#09090B", "#E4E4E7"),

        // Primary Buttons (High contrast clean)
        ["PrimaryBtnBg"] = ("#09090B", "#FFFFFF"),
        ["PrimaryBtnHover"] = ("#27272A", "#E4E4E7"),
        ["PrimaryBtnFg"] = ("#FFFFFF", "#000000"),

        // Secondary Buttons
        ["SecondaryBtnBg"] = ("#F4F4F5", "#1A1A1A"),
        ["SecondaryBtnHover"] = ("#E4E4E7", "#262626"),
        ["SecondaryBtnFg"] = ("#18181B", "#F4F4F5"),
        ["SecondaryBtnBorder"] = ("#D4D4D8", "#2E2E2E"),

        // Input Controls
        ["TextBoxBg"] = ("#FFFFFF", "#0F0F0F"),
        ["TextBoxBorder"] = ("#D4D4D8", "#262626"),
        ["TextBoxFocusBorder"] = ("#18181B", "#52525B"),
        ["ControlBorder"] = ("#D4D4D8", "#262626"),

        // Navigation
        ["NavActiveBg"] = ("#E4E4E7", "#242424"),
        ["NavActiveBorder"] = ("#18181B", "#52525B"),
        ["NavActiveFg"] = ("#09090B", "#FFFFFF"),
        ["NavInactiveFg"] = ("#71717A", "#A1A1AA"),
        ["NavHoverBg"] = ("#F4F4F5", "#1A1A1A"),

        // Workspace Preview Area - Canvas & Paper Sheet
        ["PreviewBg"] = ("#F4F4F5", "#0E0E0E"),
        ["PreviewBorder"] = ("#E4E4E7", "#1E1E1E"),
        ["PreviewSheetBg"] = ("#FFFFFF", "#FFFFFF"), // Pure white paper canvas so black barcode is always visible
        ["PreviewSheetBorder"] = ("#D4D4D8", "#333333"),

        // Code Inspector
        ["CodeBlockBg"] = ("#18181B", "#0D0D0D"),
        ["CodeBlockBorder"] = ("#27272A", "#222222"),
        ["CodeBlockFg"] = ("#0284C7", "#38BDF8"),

        // Status & Alerts
        ["StatusBg"] = ("#F4F4F5", "#0F0F0F"),
        ["StatusFg"] = ("#52525B", "#A1A1AA"),
        ["StatusSuccess"] = ("#16A34A", "#22C55E"),
        ["StatusWarning"] = ("#D97706", "#F59E0B"),
        ["StatusError"] = ("#DC2626", "#EF4444"),
        ["StatusInfo"] = ("#18181B", "#FFFFFF"),

        // Legacy compat keys
        ["HeaderCardBg"] = ("#18181B", "#18181B"),
        ["HeaderText"] = ("#FFFFFF", "#FFFFFF"),
        ["HeaderSub"] = ("#E4E4E7", "#A1A1AA"),
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

        return true;
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
