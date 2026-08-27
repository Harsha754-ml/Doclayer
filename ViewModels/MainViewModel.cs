using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using WordBarcodeStudio.Models;
using WordBarcodeStudio.Services;
using MessageBox = System.Windows.MessageBox;

namespace WordBarcodeStudio.ViewModels;

public class MainViewModel : INotifyPropertyChanged, IDisposable
{
    private readonly WordService _word = new();

    public MainViewModel()
    {
        Settings = new BarcodeSettings
        {
            Data = "https://example.com",
            BarcodeType = "QR"
        };

        BarcodeTypes = BarcodeFieldService.BarcodeTypes;

        PropertyChanged += (_, e) =>
        {
            if (e.PropertyName == nameof(SelectedBarcodeType))
                RebuildOptions();
        };

        RebuildOptions();

        GenerateCommand = new RelayCommand(_ => _ = GenerateAsync());
        OpenCommand = new RelayCommand(_ => OpenInWord());
        SaveDocxCommand = new RelayCommand(_ => SaveDocx());
        ExportPdfCommand = new RelayCommand(_ => ExportPdf());
        ToggleThemeCommand = new RelayCommand(_ => ToggleTheme());

        ThemeButtonLabel = ThemeManager.IsDark ? "Light mode" : "Dark mode";
    }

    public ICommand ToggleThemeCommand { get; }

    private string _themeButtonLabel = "Dark mode";
    public string ThemeButtonLabel
    {
        get => _themeButtonLabel;
        set { _themeButtonLabel = value; OnPropertyChanged(); }
    }

    private void ToggleTheme()
    {
        ThemeManager.Toggle();
        ThemeButtonLabel = ThemeManager.IsDark ? "Light mode" : "Dark mode";
    }

    public BarcodeSettings Settings { get; }
    public IEnumerable<string> BarcodeTypes { get; }

    public string SelectedBarcodeType
    {
        get => Settings.BarcodeType;
        set { Settings.BarcodeType = value; OnPropertyChanged(); }
    }

    public ObservableCollection<BarcodeOption> BasicOptions { get; private set; } = new();
    public ObservableCollection<BarcodeOption> AdvancedOptions { get; private set; } = new();

    private void RebuildOptions()
    {
        var all = BarcodeOptionCatalog.ForType(Settings.BarcodeType);
        BasicOptions = new ObservableCollection<BarcodeOption>(all.Where(o => o.Group == "Basic"));
        AdvancedOptions = new ObservableCollection<BarcodeOption>(all.Where(o => o.Group == "Advanced"));
        OnPropertyChanged(nameof(BasicOptions));
        OnPropertyChanged(nameof(AdvancedOptions));
    }

    private string _wordStatus = "Disconnected";
    public string WordStatus
    {
        get => _wordStatus;
        set { _wordStatus = value; OnPropertyChanged(); }
    }

    private string _generationStatus = "Ready";
    public string GenerationStatus
    {
        get => _generationStatus;
        set { _generationStatus = value; OnPropertyChanged(); }
    }

    private BitmapSource? _previewImage;
    public BitmapSource? PreviewImage
    {
        get => _previewImage;
        set { _previewImage = value; OnPropertyChanged(); OnPropertyChanged(nameof(HasPreview)); }
    }

    public bool HasPreview => _previewImage != null;

    private string _fieldCodePreview = string.Empty;
    public string FieldCodePreview
    {
        get => _fieldCodePreview;
        set { _fieldCodePreview = value; OnPropertyChanged(); }
    }

    private bool _canAct;
    public bool CanAct
    {
        get => _canAct;
        set { _canAct = value; OnPropertyChanged(); }
    }

    public ICommand GenerateCommand { get; }
    public ICommand OpenCommand { get; }
    public ICommand SaveDocxCommand { get; }
    public ICommand ExportPdfCommand { get; }

    private async Task GenerateAsync()
    {
        CanAct = false;
        PreviewImage = null;
        FieldCodePreview = string.Empty;

        try
        {
            if (!BarcodeFieldService.ValidateData(Settings.Data, Settings.BarcodeType, out string err))
            {
                GenerationStatus = "Ready";
                MessageBox.Show(err, "Word Barcode Studio", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            IEnumerable<BarcodeOption> options = BasicOptions.Concat(AdvancedOptions);
            string fieldCode = BarcodeFieldService.BuildFieldCode(Settings.Data, Settings.BarcodeType, options);

            // Runs on the UI (STA) thread so Word automation and clipboard preview work.
            GenerationStatus = "Starting Word...";
            WordStatus = "Connecting";
            await Task.Delay(30);

            _word.EnsureWord(Settings.RunWordInBackground);

            GenerationStatus = "Creating document...";
            await Task.Delay(30);

            var result = _word.GenerateBarcode(fieldCode, Settings.RunWordInBackground);

            GenerationStatus = "Inserting barcode field...";
            await Task.Delay(30);
            GenerationStatus = "Updating barcode...";
            await Task.Delay(30);

            FieldCodePreview = result.FieldCode;

            GenerationStatus = "Generating preview...";
            await Task.Delay(30);

            string? previewPath = await _word.ExtractPreviewAsync();
            if (!string.IsNullOrEmpty(previewPath) && File.Exists(previewPath))
            {
                var bmp = new BitmapImage();
                bmp.BeginInit();
                bmp.CacheOption = BitmapCacheOption.OnLoad;
                bmp.UriSource = new Uri(previewPath);
                bmp.EndInit();
                PreviewImage = bmp;
            }

            GenerationStatus = "Completed";
            WordStatus = "Connected";
        }
        catch (WordNotAvailableException ex)
        {
            GenerationStatus = "Generation failed: Microsoft Word is not available.";
            WordStatus = "Disconnected";
            MessageBox.Show(ex.Message, "Word Barcode Studio", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        catch (WordAutomationException ex)
        {
            GenerationStatus = "Generation failed.";
            MessageBox.Show(ex.Message, "Word Barcode Studio", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        catch (Exception ex)
        {
            GenerationStatus = "Generation failed: unexpected error.";
            MessageBox.Show("An unexpected error occurred.\n" + ex.Message,
                "Word Barcode Studio", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            CanAct = _word.IsWordRunning && !string.IsNullOrEmpty(FieldCodePreview);
        }
    }

    private void OpenInWord()
    {
        try { _word.OpenInWord(); }
        catch (WordAutomationException ex)
        {
            MessageBox.Show(ex.Message, "Word Barcode Studio", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void SaveDocx()
    {
        if (!_word.IsWordRunning)
        {
            MessageBox.Show("Generate a barcode first.", "Word Barcode Studio", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var dlg = new Microsoft.Win32.SaveFileDialog
        {
            FileName = "Barcode_Output.docx",
            Filter = "Word Document (*.docx)|*.docx"
        };

        if (dlg.ShowDialog() == true)
        {
            try
            {
                _word.SaveDocx(dlg.FileName);
                MessageBox.Show("Saved: " + dlg.FileName, "Word Barcode Studio",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (WordAutomationException ex)
            {
                MessageBox.Show(ex.Message, "Word Barcode Studio", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    private void ExportPdf()
    {
        if (!_word.IsWordRunning)
        {
            MessageBox.Show("Generate a barcode first.", "Word Barcode Studio", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var dlg = new Microsoft.Win32.SaveFileDialog
        {
            FileName = "Barcode_Output.pdf",
            Filter = "PDF (*.pdf)|*.pdf"
        };

        if (dlg.ShowDialog() == true)
        {
            try
            {
                _word.ExportPdf(dlg.FileName);
                MessageBox.Show("Exported: " + dlg.FileName, "Word Barcode Studio",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (WordAutomationException ex)
            {
                MessageBox.Show(ex.Message, "Word Barcode Studio", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    public void Dispose()
    {
        _word.Dispose();
        GC.SuppressFinalize(this);
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    private void OnPropertyChanged([System.Runtime.CompilerServices.CallerMemberName] string? n = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(n));
}

public class RelayCommand : ICommand
{
    private readonly Action<object?> _execute;
    public RelayCommand(Action<object?> execute) => _execute = execute;
    public event EventHandler? CanExecuteChanged { add { } remove { } }
    public bool CanExecute(object? parameter) => true;
    public void Execute(object? parameter) => _execute(parameter);
}
