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
            BarcodeType = "QR",
            RunWordInBackground = true
        };

        BarcodeTypes = BarcodeFieldService.BarcodeTypes;

        // Initialize default multi-item list
        Entries = new ObservableCollection<BarcodeEntry>
        {
            new BarcodeEntry { Label = "Product URL", Data = "https://example.com", BarcodeType = "QR", IsSelected = true },
            new BarcodeEntry { Label = "Serial Number", Data = "DOC-2026-X89", BarcodeType = "CODE128", IsSelected = true },
            new BarcodeEntry { Label = "EAN Retail Code", Data = "5901234123457", BarcodeType = "EAN13", IsSelected = true }
        };

        SelectedEntry = Entries[0];

        foreach (var entry in Entries)
        {
            entry.PropertyChanged += (s, e) =>
            {
                if (e.PropertyName == nameof(BarcodeEntry.IsSelected))
                {
                    OnPropertyChanged(nameof(SelectedEntriesCount));
                    OnPropertyChanged(nameof(GenerateButtonLabel));
                }
            };
        }

        PropertyChanged += (_, e) =>
        {
            if (e.PropertyName == nameof(SelectedBarcodeType))
                RebuildOptions();
        };

        RebuildOptions();

        // Commands
        GenerateCommand = new RelayCommand(_ => _ = GenerateAsync());
        OpenCommand = new RelayCommand(_ => OpenInWord());
        SaveDocxCommand = new RelayCommand(_ => SaveDocx());
        ExportPdfCommand = new RelayCommand(_ => ExportPdf());
        ToggleThemeCommand = new RelayCommand(_ => ToggleTheme());
        NavigateCommand = new RelayCommand(p => Navigate(p?.ToString() ?? "Generate"));
        CopyFieldCodeCommand = new RelayCommand(_ => _ = CopyFieldCodeAsync());
        ClearHistoryCommand = new RelayCommand(_ => HistoryItems.Clear());
        SelectHistoryItemCommand = new RelayCommand(p => SelectHistoryItem(p as GenerationHistoryItem));
        OpenTempFolderCommand = new RelayCommand(_ => OpenTempFolder());
        FillExampleDataCommand = new RelayCommand(_ => FillExampleData());

        // Multi-entry queue commands
        AddEntryCommand = new RelayCommand(_ => AddEntry());
        RemoveEntryCommand = new RelayCommand(p => RemoveEntry(p as BarcodeEntry));
        SelectAllEntriesCommand = new RelayCommand(_ => SetAllEntriesSelected(true));
        DeselectAllEntriesCommand = new RelayCommand(_ => SetAllEntriesSelected(false));
        SelectEntryForEditCommand = new RelayCommand(p => SelectEntryForEdit(p as BarcodeEntry));

        ThemeButtonLabel = ThemeManager.IsDark ? "Light mode" : "Dark mode";
    }

    #region Multi-Barcode Queue Management

    public ObservableCollection<BarcodeEntry> Entries { get; }

    private BarcodeEntry? _selectedEntry;
    public BarcodeEntry? SelectedEntry
    {
        get => _selectedEntry;
        set
        {
            if (_selectedEntry != value)
            {
                _selectedEntry = value;
                OnPropertyChanged();
                if (_selectedEntry != null)
                {
                    Settings.Data = _selectedEntry.Data;
                    SelectedBarcodeType = _selectedEntry.BarcodeType;
                    OnPropertyChanged(nameof(Settings));
                }
            }
        }
    }

    public int SelectedEntriesCount => Entries.Count(e => e.IsSelected);

    public ICommand AddEntryCommand { get; }
    public ICommand RemoveEntryCommand { get; }
    public ICommand SelectAllEntriesCommand { get; }
    public ICommand DeselectAllEntriesCommand { get; }
    public ICommand SelectEntryForEditCommand { get; }

    private void AddEntry()
    {
        var type = BarcodeTypes.ElementAtOrDefault((Entries.Count) % BarcodeTypes.Count()) ?? "QR";
        var info = BarcodeFieldService.GetTypeInfo(type);
        var newEntry = new BarcodeEntry
        {
            Label = $"Item #{Entries.Count + 1}",
            BarcodeType = type,
            Data = info.Example,
            IsSelected = true
        };

        newEntry.PropertyChanged += (s, e) =>
        {
            if (e.PropertyName == nameof(BarcodeEntry.IsSelected))
            {
                OnPropertyChanged(nameof(SelectedEntriesCount));
                OnPropertyChanged(nameof(GenerateButtonLabel));
            }
        };

        Entries.Add(newEntry);
        SelectedEntry = newEntry;
        OnPropertyChanged(nameof(SelectedEntriesCount));
        OnPropertyChanged(nameof(GenerateButtonLabel));
    }

    private void RemoveEntry(BarcodeEntry? entry)
    {
        if (entry != null && Entries.Count > 1)
        {
            Entries.Remove(entry);
            if (SelectedEntry == entry)
            {
                SelectedEntry = Entries.FirstOrDefault();
            }
            OnPropertyChanged(nameof(SelectedEntriesCount));
            OnPropertyChanged(nameof(GenerateButtonLabel));
        }
    }

    private void SetAllEntriesSelected(bool selected)
    {
        foreach (var entry in Entries)
        {
            entry.IsSelected = selected;
        }
        OnPropertyChanged(nameof(SelectedEntriesCount));
        OnPropertyChanged(nameof(GenerateButtonLabel));
    }

    private void SelectEntryForEdit(BarcodeEntry? entry)
    {
        if (entry != null)
        {
            SelectedEntry = entry;
        }
    }

    #endregion

    #region Navigation

    private string _currentTab = "Generate";
    public string CurrentTab
    {
        get => _currentTab;
        set
        {
            if (_currentTab != value)
            {
                _currentTab = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsGenerateTabActive));
                OnPropertyChanged(nameof(IsHistoryTabActive));
                OnPropertyChanged(nameof(IsTemplatesTabActive));
                OnPropertyChanged(nameof(IsSettingsTabActive));
                OnPropertyChanged(nameof(IsAboutTabActive));
            }
        }
    }

    public bool IsGenerateTabActive => CurrentTab == "Generate";
    public bool IsHistoryTabActive => CurrentTab == "History";
    public bool IsTemplatesTabActive => CurrentTab == "Templates";
    public bool IsSettingsTabActive => CurrentTab == "Settings";
    public bool IsAboutTabActive => CurrentTab == "About";

    public ICommand NavigateCommand { get; }

    public void Navigate(string tab)
    {
        CurrentTab = tab;
    }

    #endregion

    #region Theme & Meta

    public ICommand ToggleThemeCommand { get; }

    private string _themeButtonLabel = "Dark mode";
    public string ThemeButtonLabel
    {
        get => _themeButtonLabel;
        set { _themeButtonLabel = value; OnPropertyChanged(); }
    }

    public bool IsDarkMode => ThemeManager.IsDark;

    private void ToggleTheme()
    {
        ThemeManager.Toggle();
        ThemeButtonLabel = ThemeManager.IsDark ? "Light mode" : "Dark mode";
        OnPropertyChanged(nameof(IsDarkMode));
    }

    public string AppVersion => "v0.1.0";
    public string WordTempDirectory => _word.TempDirectory;

    public ICommand OpenTempFolderCommand { get; }
    private void OpenTempFolder()
    {
        try
        {
            if (Directory.Exists(_word.TempDirectory))
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = _word.TempDirectory,
                    UseShellExecute = true
                });
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "DocLayer", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    #endregion

    #region Barcode Settings & Options

    public BarcodeSettings Settings { get; }
    public IEnumerable<string> BarcodeTypes { get; }

    public BarcodeTypeInfo CurrentTypeInfo => BarcodeFieldService.GetTypeInfo(Settings.BarcodeType);

    public ICommand FillExampleDataCommand { get; }

    public void FillExampleData()
    {
        Settings.Data = CurrentTypeInfo.Example;
        if (SelectedEntry != null)
        {
            SelectedEntry.Data = Settings.Data;
        }
        OnPropertyChanged(nameof(Settings));
    }

    public string SelectedBarcodeType
    {
        get => Settings.BarcodeType;
        set
        {
            if (Settings.BarcodeType != value)
            {
                Settings.BarcodeType = value;
                if (SelectedEntry != null)
                {
                    SelectedEntry.BarcodeType = value;
                }
                OnPropertyChanged();
                OnPropertyChanged(nameof(CurrentTypeInfo));
            }
        }
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

    #endregion

    #region State & Status

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

    private bool _isGenerating;
    public bool IsGenerating
    {
        get => _isGenerating;
        set
        {
            _isGenerating = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(GenerateButtonLabel));
        }
    }

    private string _generateButtonLabel = "GENERATE";
    public string GenerateButtonLabel
    {
        get
        {
            if (_isGenerating) return "GENERATING...";
            int count = SelectedEntriesCount;
            if (count > 1) return $"GENERATE ({count} ITEMS)";
            if (count == 1) return "GENERATE (1 ITEM)";
            return "GENERATE";
        }
        set { _generateButtonLabel = value; OnPropertyChanged(); }
    }

    private BitmapSource? _previewImage;
    public BitmapSource? PreviewImage
    {
        get => _previewImage;
        set
        {
            _previewImage = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(HasPreview));
        }
    }

    public bool HasPreview => _previewImage != null;

    private string _fieldCodePreview = string.Empty;
    public string FieldCodePreview
    {
        get => _fieldCodePreview;
        set { _fieldCodePreview = value; OnPropertyChanged(); }
    }

    private string _copyFeedbackText = "Copy";
    public string CopyFeedbackText
    {
        get => _copyFeedbackText;
        set { _copyFeedbackText = value; OnPropertyChanged(); }
    }

    private bool _canAct;
    public bool CanAct
    {
        get => _canAct;
        set { _canAct = value; OnPropertyChanged(); }
    }

    public ObservableCollection<GenerationHistoryItem> HistoryItems { get; } = new();

    public ICommand ClearHistoryCommand { get; }
    public ICommand SelectHistoryItemCommand { get; }

    private void SelectHistoryItem(GenerationHistoryItem? item)
    {
        if (item == null) return;
        Settings.Data = item.Data;
        OnPropertyChanged(nameof(Settings));
        SelectedBarcodeType = item.BarcodeType;
        FieldCodePreview = item.FieldCode;
        PreviewImage = item.PreviewImage;
        CurrentTab = "Generate";
    }

    #endregion

    #region Commands & Generation Workflow

    public ICommand GenerateCommand { get; }
    public ICommand OpenCommand { get; }
    public ICommand SaveDocxCommand { get; }
    public ICommand ExportPdfCommand { get; }
    public ICommand CopyFieldCodeCommand { get; }

    private async Task CopyFieldCodeAsync()
    {
        if (string.IsNullOrWhiteSpace(FieldCodePreview)) return;
        try
        {
            System.Windows.Clipboard.SetText(FieldCodePreview);
            CopyFeedbackText = "Copied!";
            await Task.Delay(2000);
            CopyFeedbackText = "Copy";
        }
        catch
        {
            CopyFeedbackText = "Failed";
        }
    }

    private async Task GenerateAsync()
    {
        if (IsGenerating) return;

        // Sync currently edited data back to SelectedEntry
        if (SelectedEntry != null)
        {
            SelectedEntry.Data = Settings.Data;
            SelectedEntry.BarcodeType = Settings.BarcodeType;
        }

        var selectedEntries = Entries.Where(e => e.IsSelected).ToList();
        if (selectedEntries.Count == 0)
        {
            MessageBox.Show("Please select at least one barcode to generate.", "DocLayer", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        // Validate all selected entries
        for (int i = 0; i < selectedEntries.Count; i++)
        {
            var entry = selectedEntries[i];
            if (!BarcodeFieldService.ValidateData(entry.Data, entry.BarcodeType, out string err))
            {
                GenerationStatus = "Ready";
                GenerateButtonLabel = "GENERATE";
                MessageBox.Show($"Validation error in [{entry.Label}]:\n\n{err}", "DocLayer", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
        }

        IsGenerating = true;
        CanAct = false;
        PreviewImage = null;
        FieldCodePreview = string.Empty;

        try
        {
            GenerationStatus = "Starting Word...";
            WordStatus = "Connecting";
            await Task.Delay(30);

            _word.EnsureWord(Settings.RunWordInBackground);

            GenerationStatus = "Creating document...";
            await Task.Delay(30);

            IEnumerable<BarcodeOption> options = BasicOptions.Concat(AdvancedOptions);
            var buildItems = new List<(string FieldCode, string Label)>();

            foreach (var entry in selectedEntries)
            {
                // If it's the currently selected entry, use full options from UI
                IEnumerable<BarcodeOption> itemOptions = (entry == SelectedEntry)
                    ? options
                    : BarcodeOptionCatalog.ForType(entry.BarcodeType);

                string code = BarcodeFieldService.BuildFieldCode(entry.Data, entry.BarcodeType, itemOptions);
                buildItems.Add((code, entry.Label));
            }

            GenerationStatus = $"Generating {selectedEntries.Count} barcode(s)...";
            await Task.Delay(30);

            BarcodeResult result;
            if (buildItems.Count == 1)
            {
                result = _word.GenerateBarcode(buildItems[0].FieldCode, Settings.RunWordInBackground);
            }
            else
            {
                result = _word.GenerateMultipleBarcodes(buildItems, Settings.RunWordInBackground);
            }

            FieldCodePreview = result.FieldCode;

            GenerationStatus = "Generating live preview...";
            await Task.Delay(50);

            string? previewPath = await _word.ExtractPreviewAsync();
            if (!string.IsNullOrEmpty(previewPath) && File.Exists(previewPath))
            {
                var bmp = new BitmapImage();
                bmp.BeginInit();
                bmp.CacheOption = BitmapCacheOption.OnLoad;
                bmp.UriSource = new Uri(previewPath);
                bmp.EndInit();
                bmp.Freeze();
                PreviewImage = bmp;
            }

            GenerationStatus = "Completed";
            WordStatus = "Connected";
            GenerateButtonLabel = "GENERATED";

            // Add to session history
            var historyItem = new GenerationHistoryItem
            {
                BarcodeType = selectedEntries.Count == 1 ? selectedEntries[0].BarcodeType : $"BATCH ({selectedEntries.Count})",
                Data = string.Join("; ", selectedEntries.Select(e => $"{e.Label}: {e.Data}")),
                FieldCode = result.FieldCode,
                DocxPath = result.DocxPath,
                PreviewImage = PreviewImage
            };
            HistoryItems.Insert(0, historyItem);

            _ = Task.Run(async () =>
            {
                await Task.Delay(2500);
                GenerateButtonLabel = "GENERATE";
            });
        }
        catch (WordNotAvailableException ex)
        {
            GenerationStatus = "Failed: Microsoft Word not found";
            WordStatus = "Disconnected";
            GenerateButtonLabel = "GENERATION FAILED";
            MessageBox.Show(ex.Message, "DocLayer", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        catch (WordAutomationException ex)
        {
            GenerationStatus = "Failed: Word automation error";
            GenerateButtonLabel = "GENERATION FAILED";
            MessageBox.Show(ex.Message, "DocLayer", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        catch (Exception ex)
        {
            GenerationStatus = "Failed: unexpected error";
            GenerateButtonLabel = "GENERATION FAILED";
            MessageBox.Show("An unexpected error occurred.\n" + ex.Message,
                "DocLayer", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            IsGenerating = false;
            CanAct = _word.IsWordRunning && !string.IsNullOrEmpty(FieldCodePreview);
        }
    }

    private void OpenInWord()
    {
        try
        {
            _word.OpenInWord();
        }
        catch (WordAutomationException ex)
        {
            MessageBox.Show(ex.Message, "DocLayer", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void SaveDocx()
    {
        if (!_word.IsWordRunning)
        {
            MessageBox.Show("Generate a barcode first.", "DocLayer", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var dlg = new Microsoft.Win32.SaveFileDialog
        {
            FileName = $"DocLayer_Output_{DateTime.Now:yyyyMMdd_HHmmss}.docx",
            Filter = "Word Document (*.docx)|*.docx"
        };

        if (dlg.ShowDialog() == true)
        {
            try
            {
                _word.SaveDocx(dlg.FileName);
                MessageBox.Show("Saved: " + dlg.FileName, "DocLayer",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (WordAutomationException ex)
            {
                MessageBox.Show(ex.Message, "DocLayer", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    private void ExportPdf()
    {
        if (!_word.IsWordRunning)
        {
            MessageBox.Show("Generate a barcode first.", "DocLayer", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var dlg = new Microsoft.Win32.SaveFileDialog
        {
            FileName = $"DocLayer_Output_{DateTime.Now:yyyyMMdd_HHmmss}.pdf",
            Filter = "PDF (*.pdf)|*.pdf"
        };

        if (dlg.ShowDialog() == true)
        {
            try
            {
                _word.ExportPdf(dlg.FileName);
                MessageBox.Show("Exported: " + dlg.FileName, "DocLayer",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (WordAutomationException ex)
            {
                MessageBox.Show(ex.Message, "DocLayer", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    #endregion

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
    private readonly Predicate<object?>? _canExecute;

    public RelayCommand(Action<object?> execute, Predicate<object?>? canExecute = null)
    {
        _execute = execute ?? throw new ArgumentNullException(nameof(execute));
        _canExecute = canExecute;
    }

    public event EventHandler? CanExecuteChanged
    {
        add => CommandManager.RequerySuggested += value;
        remove => CommandManager.RequerySuggested -= value;
    }

    public bool CanExecute(object? parameter) => _canExecute == null || _canExecute(parameter);
    public void Execute(object? parameter) => _execute(parameter);
}
