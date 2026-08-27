using System.Collections.Generic;
using System.ComponentModel;

namespace WordBarcodeStudio.Models;

public enum BarcodeControlType
{
    ComboBox,
    Number,
    CheckBox
}

public class ChoiceItem
{
    public string Display { get; }
    public string Value { get; }
    public ChoiceItem(string display, string value)
    {
        Display = display;
        Value = value;
    }
}

/// <summary>
/// Describes a single DISPLAYBARCODE switch and how it should be rendered in the UI.
/// A BarcodeOption carries its own current Value so controls can bind directly to it.
/// </summary>
public class BarcodeOption : INotifyPropertyChanged
{
    public string Id { get; init; } = "";
    public string Switch { get; init; } = "";
    public string Label { get; init; } = "";
    public BarcodeControlType Control { get; init; }
    public string Group { get; init; } = "Basic";
    public bool Optional { get; init; }
    public IReadOnlyList<ChoiceItem>? Items { get; init; }
    public object? DefaultValue { get; init; }

    private object? _value;
    public object? Value
    {
        get => _value ?? DefaultValue;
        set { if (!Equals(_value, value)) { _value = value; OnPropertyChanged(); } }
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    private void OnPropertyChanged([System.Runtime.CompilerServices.CallerMemberName] string? n = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(n));
}

/// <summary>
/// Returns the set of DISPLAYBARCODE switches that are valid for a given barcode type.
/// This is the single source of truth for "all the things Word can offer" per type.
/// </summary>
public static class BarcodeOptionCatalog
{
    public static List<BarcodeOption> ForType(string type)
    {
        var opts = new List<BarcodeOption>
        {
            Scale(),
            Rotation(),
            ShowText()
        };

        if (type == "QR") opts.Add(ErrorCorrection());

        // Advanced (valid for every type)
        opts.Add(Unicode());
        opts.Add(Height());
        opts.Add(Foreground());
        opts.Add(Background());

        // Type-specific advanced switches
        if (type is "EAN13" or "EAN8" or "UPCA" or "UPCE") opts.Add(FixCheckDigit());
        if (type == "CODE39") opts.Add(StartStop());
        if (type == "ITF14") opts.Add(CaseCode());

        return opts;
    }

    private static BarcodeOption Scale() => new()
    {
        Id = "Scale", Switch = "s", Label = "Scale (%)", Control = BarcodeControlType.Number,
        DefaultValue = "100", Optional = false, Group = "Basic"
    };

    private static BarcodeOption Rotation() => new()
    {
        Id = "Rotation", Switch = "r", Label = "Rotation", Control = BarcodeControlType.ComboBox,
        Items = new[]
        {
            new ChoiceItem("0°", "0"),
            new ChoiceItem("90°", "1"),
            new ChoiceItem("180°", "2"),
            new ChoiceItem("270°", "3")
        },
        DefaultValue = "0", Optional = true, Group = "Basic"
    };

    private static BarcodeOption ShowText() => new()
    {
        Id = "ShowText", Switch = "t", Label = "Show encoded text", Control = BarcodeControlType.CheckBox,
        DefaultValue = false, Group = "Basic"
    };

    private static BarcodeOption ErrorCorrection() => new()
    {
        Id = "ErrorCorrection", Switch = "q", Label = "Error correction (QR)", Control = BarcodeControlType.ComboBox,
        Items = new[]
        {
            new ChoiceItem("L", "L"),
            new ChoiceItem("M", "M"),
            new ChoiceItem("Q", "Q"),
            new ChoiceItem("H", "H")
        },
        DefaultValue = "H", Group = "Basic"
    };

    private static BarcodeOption Unicode() => new()
    {
        Id = "Unicode", Switch = "u", Label = "Unicode data (\\u)", Control = BarcodeControlType.CheckBox,
        DefaultValue = false, Group = "Advanced"
    };

    private static BarcodeOption Height() => new()
    {
        Id = "Height", Switch = "h", Label = "Height (twips)", Control = BarcodeControlType.Number,
        DefaultValue = "", Optional = true, Group = "Advanced"
    };

    private static BarcodeOption Foreground() => new()
    {
        Id = "Foreground", Switch = "f", Label = "Foreground color", Control = BarcodeControlType.ComboBox,
        Items = Colors(), DefaultValue = "0x000000", Optional = true, Group = "Advanced"
    };

    private static BarcodeOption Background() => new()
    {
        Id = "Background", Switch = "b", Label = "Background color", Control = BarcodeControlType.ComboBox,
        Items = Colors(), DefaultValue = "0x000000", Optional = true, Group = "Advanced"
    };

    private static BarcodeOption FixCheckDigit() => new()
    {
        Id = "FixCheckDigit", Switch = "x", Label = "Fix invalid check digit (\\x)", Control = BarcodeControlType.CheckBox,
        DefaultValue = false, Group = "Advanced"
    };

    private static BarcodeOption StartStop() => new()
    {
        Id = "StartStop", Switch = "d", Label = "Add Start/Stop chars (\\d)", Control = BarcodeControlType.CheckBox,
        DefaultValue = false, Group = "Advanced"
    };

    private static BarcodeOption CaseCode() => new()
    {
        Id = "CaseCode", Switch = "c", Label = "ITF14 case code style", Control = BarcodeControlType.ComboBox,
        Items = new[]
        {
            new ChoiceItem("STD", "STD"),
            new ChoiceItem("2", "2"),
            new ChoiceItem("3", "3")
        },
        DefaultValue = "STD", Group = "Advanced"
    };

    private static IReadOnlyList<ChoiceItem> Colors() => new[]
    {
        new ChoiceItem("Black", "0x000000"),
        new ChoiceItem("White", "0xFFFFFF"),
        new ChoiceItem("Red", "0xFF0000"),
        new ChoiceItem("Green", "0x00FF00"),
        new ChoiceItem("Blue", "0x0000FF"),
        new ChoiceItem("Yellow", "0xFFFF00"),
        new ChoiceItem("Cyan", "0x00FFFF"),
        new ChoiceItem("Magenta", "0xFF00FF")
    };
}
