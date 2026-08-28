using System;
using System.Windows.Media.Imaging;

namespace WordBarcodeStudio.Models;

public class GenerationHistoryItem
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N")[..8];
    public DateTime Timestamp { get; set; } = DateTime.Now;
    public string BarcodeType { get; set; } = string.Empty;
    public string Data { get; set; } = string.Empty;
    public string FieldCode { get; set; } = string.Empty;
    public string? DocxPath { get; set; }
    public BitmapSource? PreviewImage { get; set; }

    public string TimeFormatted => Timestamp.ToString("HH:mm:ss");
    public string DateFormatted => Timestamp.ToString("MMM dd, yyyy");
    public string TruncatedData => Data.Length > 45 ? Data[..42] + "..." : Data;
}
