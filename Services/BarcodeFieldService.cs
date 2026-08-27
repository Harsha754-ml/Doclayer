using System.Collections.Generic;
using System.Linq;
using System.Text;
using WordBarcodeStudio.Models;

namespace WordBarcodeStudio.Services;

public static class BarcodeFieldService
{
    public static IReadOnlyList<string> BarcodeTypes { get; } = new[]
    {
        "QR", "CODE128", "CODE39", "EAN13", "EAN8", "UPCA", "UPCE", "ITF14"
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
            error = "Enter barcode data first.";
            return false;
        }

        switch (barcodeType)
        {
            case "EAN13":
                if (!IsDigits(data, 12, 13)) { error = Invalid(); return false; }
                break;
            case "EAN8":
                if (!IsDigits(data, 7, 8)) { error = Invalid(); return false; }
                break;
            case "UPCA":
                if (!IsDigits(data, 11, 12)) { error = Invalid(); return false; }
                break;
            case "UPCE":
                if (!IsDigits(data, 6, 8)) { error = Invalid(); return false; }
                break;
            case "ITF14":
                if (!IsDigits(data, 13, 14)) { error = Invalid(); return false; }
                break;
            // QR, CODE128 and CODE39 accept free text.
        }

        return true;

        static string Invalid() => "The provided data is not valid for the selected barcode type.";
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
}
