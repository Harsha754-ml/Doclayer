using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace WordBarcodeStudio.Models;

public class BarcodeEntry : INotifyPropertyChanged
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N")[..8];

    private bool _isSelected = true;
    public bool IsSelected
    {
        get => _isSelected;
        set { if (_isSelected != value) { _isSelected = value; OnPropertyChanged(); } }
    }

    private string _label = "Barcode Item";
    public string Label
    {
        get => _label;
        set { if (_label != value) { _label = value; OnPropertyChanged(); } }
    }

    private string _data = "https://example.com";
    public string Data
    {
        get => _data;
        set { if (_data != value) { _data = value; OnPropertyChanged(); OnPropertyChanged(nameof(TruncatedData)); } }
    }

    private string _barcodeType = "QR";
    public string BarcodeType
    {
        get => _barcodeType;
        set { if (_barcodeType != value) { _barcodeType = value; OnPropertyChanged(); } }
    }

    private int _scale = 100;
    public int Scale
    {
        get => _scale;
        set { if (_scale != value) { _scale = value; OnPropertyChanged(); } }
    }

    private string _errorCorrection = "H";
    public string ErrorCorrection
    {
        get => _errorCorrection;
        set { if (_errorCorrection != value) { _errorCorrection = value; OnPropertyChanged(); } }
    }

    private bool _showText;
    public bool ShowText
    {
        get => _showText;
        set { if (_showText != value) { _showText = value; OnPropertyChanged(); } }
    }

    public string TruncatedData => Data.Length > 28 ? Data[..25] + "..." : Data;

    public event PropertyChangedEventHandler? PropertyChanged;

    private void OnPropertyChanged([CallerMemberName] string? n = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(n));
}
