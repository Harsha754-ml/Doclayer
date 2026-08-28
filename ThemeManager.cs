using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Media;

namespace WordBarcodeStudio;

/// <summary>
/// Swaps the application's colour resources between a refined dark and light palette.
/// The XAML binds to these resources via DynamicResource, so changing them here
/// re-themes the whole UI at runtime.
/// </summary>
internal static class ThemeManager
{
    public static bool IsDark { get; private set; }

    private static readonly Dictionary<string, (string Light, string Dark)> Palette = new()
    {
        // Application surfaces
        ["AppBg"] = ("#F8FAFC", "#07111F"),
        ["SidebarBg"] = ("#FFFFFF", "#0B1527"),
        ["SidebarBorder"] = ("#E2E8F0", "#15223A"),
        ["HeaderBg"] = ("#FFFFFF", "#0A1324"),
        ["HeaderBorder"] = ("#E2E8F0", "#15223A"),
        ["CardBg"] = ("#FFFFFF", "#0F1C32"),
        ["CardBorder"] = ("#E2E8F0", "#1E2F4D"),
        ["CardBgElevated"] = ("#F8FAFC", "#14243F"),
        ["CardBorderSubtle"] = ("#F1F5F9", "#172744"),

        // Typography & Text
        ["TextPrimary"] = ("#0F172A", "#F8FAFC"),
        ["TextSecondary"] = ("#475569", "#94A3B8"),
        ["MutedText"] = ("#94A3B8", "#64748B"),
        ["FieldLabelFg"] = ("#334155", "#94A3B8"),
        ["InputText"] = ("#0F172A", "#F8FAFC"),

        // Brand & Accents
        ["PrimaryBrand"] = ("#2563EB", "#3B82F6"),
        ["AccentPurple"] = ("#7C3AED", "#8B5CF6"),
        ["AccentCyan"] = ("#0891B2", "#22D3EE"),

        // Primary Buttons
        ["PrimaryBtnBg"] = ("#2563EB", "#2563EB"),
        ["PrimaryBtnHover"] = ("#1D4ED8", "#3B82F6"),
        ["PrimaryBtnFg"] = ("#FFFFFF", "#FFFFFF"),

        // Secondary Buttons
        ["SecondaryBtnBg"] = ("#F1F5F9", "#14223A"),
        ["SecondaryBtnHover"] = ("#E2E8F0", "#1C2E4E"),
        ["SecondaryBtnFg"] = ("#1E293B", "#E2E8F0"),
        ["SecondaryBtnBorder"] = ("#CBD5E1", "#223556"),

        // Input Controls
        ["TextBoxBg"] = ("#FFFFFF", "#091222"),
        ["TextBoxBorder"] = ("#CBD5E1", "#1C2E4C"),
        ["TextBoxFocusBorder"] = ("#2563EB", "#3B82F6"),
        ["ControlBorder"] = ("#CBD5E1", "#1E2F4D"),

        // Navigation
        ["NavActiveBg"] = ("#EEF2FF", "#182846"),
        ["NavActiveBorder"] = ("#6366F1", "#3B82F6"),
        ["NavActiveFg"] = ("#2563EB", "#60A5FA"),
        ["NavInactiveFg"] = ("#64748B", "#94A3B8"),
        ["NavHoverBg"] = ("#F8FAFC", "#101D35"),

        // Workspace Preview Area
        ["PreviewBg"] = ("#F1F5F9", "#060D18"),
        ["PreviewBorder"] = ("#E2E8F0", "#162339"),
        ["PreviewSheetBg"] = ("#FFFFFF", "#0D1829"),
        ["PreviewSheetBorder"] = ("#CBD5E1", "#1C2D49"),

        // Code Inspector
        ["CodeBlockBg"] = ("#0F172A", "#050B14"),
        ["CodeBlockBorder"] = ("#1E293B", "#142036"),
        ["CodeBlockFg"] = ("#0284C7", "#38BDF8"),

        // Status & Alerts
        ["StatusBg"] = ("#F8FAFC", "#070E1A"),
        ["StatusFg"] = ("#334155", "#94A3B8"),
        ["StatusSuccess"] = ("#16A34A", "#22C55E"),
        ["StatusWarning"] = ("#D97706", "#F59E0B"),
        ["StatusError"] = ("#DC2626", "#EF4444"),
        ["StatusInfo"] = ("#2563EB", "#3B82F6"),

        // Legacy compat keys
        ["HeaderCardBg"] = ("#2563EB", "#1E3A8A"),
        ["HeaderText"] = ("#FFFFFF", "#FFFFFF"),
        ["HeaderSub"] = ("#DBEAFE", "#BFDBFE"),
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

        // Default to dark mode for our modern developer-first UI
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
