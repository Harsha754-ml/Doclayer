using System.Collections.Generic;
using System.Linq;
using System.Text;
using WordBarcodeStudio.Models;

namespace WordBarcodeStudio.Services;

public record BarcodeTypeInfo(string Type, string Name, string ExpectedFormat, string Example, string Description);

public static class BarcodeFieldService
{
    public static IReadOnlyList<string> BarcodeTypes { get; } = new[]
    {
        "QR", "CODE128", "CODE39", "EAN13", "EAN8", "UPCA", "UPCE", "ITF14"
    };

    public static BarcodeTypeInfo GetTypeInfo(string barcodeType) => barcodeType.ToUpperInvariant() switch
    {
        "QR" => new BarcodeTypeInfo(
            "QR",
            "QR Code (2D)",
            "Any text, URL, or data payload",
            "https://example.com",
            "High-capacity 2D matrix code supporting full text, URLs, and alphanumeric data with error correction."),

        "CODE128" => new BarcodeTypeInfo(
            "CODE128",
            "Code 128 (1D)",
            "Any standard ASCII text (letters, numbers, symbols)",
            "DOC-2026-X89",
            "High-density 1D barcode encoding the full 128 ASCII character set."),

        "CODE39" => new BarcodeTypeInfo(
            "CODE39",
            "Code 39 (1D)",
            "Uppercase letters (A-Z), digits (0-9), and characters (- . $ / + % space)",
            "PART-9872",
            "Standard industrial 1D barcode supporting uppercase alphanumeric data."),

        "EAN13" => new BarcodeTypeInfo(
            "EAN13",
            "EAN-13 (International Article Number)",
            "12 or 13 numeric digits only",
            "5901234123457",
            "13-digit retail barcode standard. Enter 12 digits (Word auto-calculates check digit) or all 13 digits."),

        "EAN8" => new BarcodeTypeInfo(
            "EAN8",
            "EAN-8 (Short Product Code)",
            "7 or 8 numeric digits only",
            "96385074",
            "Compact 8-digit retail barcode for small packages. Enter 7 or 8 digits."),

        "UPCA" => new BarcodeTypeInfo(
            "UPCA",
            "UPC-A (Universal Product Code)",
            "11 or 12 numeric digits only",
            "012345678905",
            "12-digit standard product barcode. Enter 11 digits or full 12 digits."),

        "UPCE" => new BarcodeTypeInfo(
            "UPCE",
            "UPC-E (Zero-Suppressed UPC)",
            "6, 7, or 8 numeric digits only",
            "01234565",
            "Zero-suppressed compact version of UPC-A for small packages."),

        "ITF14" => new BarcodeTypeInfo(
            "ITF14",
            "ITF-14 (Interleaved 2 of 5)",
            "13 or 14 numeric digits only",
            "10012345678902",
            "14-digit carton/shipping container packaging barcode."),

        _ => new BarcodeTypeInfo(
            barcodeType,
            barcodeType,
            "Valid data for " + barcodeType,
            "123456",
            "Word DISPLAYBARCODE format.")
    };

    /// <summary>
    /// Builds the Word field code from the data, the selected barcode type, and the
    /// applicable switch options. Only options relevant to the selected type are passed in.
    /// </summary>
    public static string BuildFieldCode(string data, string barcodeType, IEnumerable<BarcodeOption> options)
    {
        var sb = new StringBuilder();
        sb.Append(" DISPLAYBARCODE ");
        sb.Append('"').Append(EscapeData(data)).Append('"');
        sb.Append(' ').Append(barcodeType);

        foreach (var o in options)
        {
            if (o.Control == BarcodeControlType.CheckBox)
            {
                // Flag switch: no argument, just present when enabled.
                if (o.Value is bool b && b)
                {
                    sb.Append(" \\").Append(o.Switch);
                }
                continue;
            }

            var val = (o.Value?.ToString() ?? "").Trim();
            if (o.Optional)
            {
                if (string.IsNullOrWhiteSpace(val)) continue;
                if (o.Control == BarcodeControlType.Number &&
                    int.TryParse(val, out int n) && n <= 0) continue;
            }

            if (string.IsNullOrWhiteSpace(val)) continue;

            sb.Append(" \\").Append(o.Switch).Append(' ').Append(val);
        }

        return sb.ToString();
    }

    private static string EscapeData(string data)
    {
        // In Word field codes, a literal double quote inside a quoted string is
        // escaped by doubling it.
        return data.Replace("\"", "\"\"");
    }

    public static bool ValidateData(string data, string barcodeType, out string error)
    {
        error = string.Empty;

        if (string.IsNullOrWhiteSpace(data))
        {
            var info = GetTypeInfo(barcodeType);
            error = $"Please enter data to encode.\n\nExpected for {info.Name}:\n{info.ExpectedFormat}\n\nExample:\n{info.Example}";
            return false;
        }

        data = data.Trim();
        var typeInfo = GetTypeInfo(barcodeType);

        switch (barcodeType.ToUpperInvariant())
        {
            case "EAN13":
                if (!IsDigits(data, 12, 13))
                {
                    error = $"Invalid data for EAN-13.\n\nEAN-13 requires exactly 12 or 13 numeric digits.\n\nYour input: \"{data}\" ({data.Length} characters)\nExpected Example: {typeInfo.Example}";
                    return false;
                }
                break;

            case "EAN8":
                if (!IsDigits(data, 7, 8))
                {
                    error = $"Invalid data for EAN-8.\n\nEAN-8 requires exactly 7 or 8 numeric digits.\n\nYour input: \"{data}\" ({data.Length} characters)\nExpected Example: {typeInfo.Example}";
                    return false;
                }
                break;

            case "UPCA":
                if (!IsDigits(data, 11, 12))
                {
                    error = $"Invalid data for UPC-A.\n\nUPC-A requires exactly 11 or 12 numeric digits.\n\nYour input: \"{data}\" ({data.Length} characters)\nExpected Example: {typeInfo.Example}";
                    return false;
                }
                break;

            case "UPCE":
                if (!IsDigits(data, 6, 8))
                {
                    error = $"Invalid data for UPC-E.\n\nUPC-E requires 6, 7, or 8 numeric digits.\n\nYour input: \"{data}\" ({data.Length} characters)\nExpected Example: {typeInfo.Example}";
                    return false;
                }
                break;

            case "ITF14":
                if (!IsDigits(data, 13, 14))
                {
                    error = $"Invalid data for ITF-14.\n\nITF-14 requires exactly 13 or 14 numeric digits.\n\nYour input: \"{data}\" ({data.Length} characters)\nExpected Example: {typeInfo.Example}";
                    return false;
                }
                break;

            case "CODE39":
                if (!IsValidCode39(data))
                {
                    error = $"Invalid characters for Code 39.\n\nCode 39 only supports uppercase letters (A-Z), numbers (0-9), and symbols (- . $ / + % space).\n\nYour input: \"{data}\"\nExpected Example: {typeInfo.Example}";
                    return false;
                }
                break;

            // QR and CODE128 accept general text.
        }

        return true;
    }

    private static bool IsDigits(string s, int min, int max)
    {
        if (s.Length < min || s.Length > max) return false;
        foreach (var c in s)
        {
            if (!char.IsDigit(c)) return false;
        }
        return true;
    }

    private static bool IsValidCode39(string s)
    {
        const string validChars = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-. $/+%";
        foreach (var c in s.ToUpperInvariant())
        {
            if (!validChars.Contains(c)) return false;
        }
        return true;
    }
}
