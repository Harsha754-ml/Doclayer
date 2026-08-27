namespace WordBarcodeStudio.Models;

public class BarcodeSettings
{
    public string Data { get; set; } = string.Empty;
    public string BarcodeType { get; set; } = "QR";
    public bool RunWordInBackground { get; set; } = true;
}
