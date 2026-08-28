using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using DocLayer.Setup.Services;

namespace DocLayer.Setup.ViewModels;

public enum InstallerStep
{
    Welcome = 0,
    Eula = 1,
    Destination = 2,
    Diagnostics = 3,
    Installing = 4,
    Completed = 5,
    Uninstall = 6,
    Uninstalled = 7
}

public class InstallerViewModel : INotifyPropertyChanged
{
    private readonly InstallService _installer = new();
    private readonly bool _isUninstallMode;

    public InstallerViewModel(bool isUninstall = false)
    {
        _isUninstallMode = isUninstall;

        if (isUninstall)
        {
            CurrentStep = InstallerStep.Uninstall;
        }
        else
        {
            CurrentStep = InstallerStep.Welcome;
        }

        InstallPath = InstallService.DefaultInstallPath;

        // Initialize Commands
        NextStepCommand = new RelayCommand(_ => NextStep(), _ => CanGoNext());
        PrevStepCommand = new RelayCommand(_ => PrevStep(), _ => CanGoPrev());
        CancelCommand = new RelayCommand(_ => System.Windows.Application.Current.Shutdown());
        BrowsePathCommand = new RelayCommand(_ => BrowseDestination());
        FinishCommand = new RelayCommand(_ => Finish());
        StartUninstallCommand = new RelayCommand(_ => _ = RunUninstallAsync());

        RefreshDiagnostics();
    }

    #region Wizard Steps & Navigation

    private InstallerStep _currentStep;
    public InstallerStep CurrentStep
    {
        get => _currentStep;
        set
        {
            if (_currentStep != value)
            {
                _currentStep = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsWelcomeStep));
                OnPropertyChanged(nameof(IsEulaStep));
                OnPropertyChanged(nameof(IsDestinationStep));
                OnPropertyChanged(nameof(IsDiagnosticsStep));
                OnPropertyChanged(nameof(IsInstallingStep));
                OnPropertyChanged(nameof(IsCompletedStep));
                OnPropertyChanged(nameof(IsUninstallStep));
                OnPropertyChanged(nameof(IsUninstalledStep));
                OnPropertyChanged(nameof(NextButtonText));
                OnPropertyChanged(nameof(CanShowBackButton));
                OnPropertyChanged(nameof(CanShowCancelButton));
                OnPropertyChanged(nameof(CanShowNextButton));
                OnPropertyChanged(nameof(CanShowFinishButton));
            }
        }
    }

    public bool IsWelcomeStep => CurrentStep == InstallerStep.Welcome;
    public bool IsEulaStep => CurrentStep == InstallerStep.Eula;
    public bool IsDestinationStep => CurrentStep == InstallerStep.Destination;
    public bool IsDiagnosticsStep => CurrentStep == InstallerStep.Diagnostics;
    public bool IsInstallingStep => CurrentStep == InstallerStep.Installing;
    public bool IsCompletedStep => CurrentStep == InstallerStep.Completed;
    public bool IsUninstallStep => CurrentStep == InstallerStep.Uninstall;
    public bool IsUninstalledStep => CurrentStep == InstallerStep.Uninstalled;

    public bool CanShowBackButton => CurrentStep is InstallerStep.Eula or InstallerStep.Destination or InstallerStep.Diagnostics;
    public bool CanShowCancelButton => CurrentStep is not InstallerStep.Installing and not InstallerStep.Completed and not InstallerStep.Uninstalled;
    public bool CanShowNextButton => CurrentStep is InstallerStep.Welcome or InstallerStep.Eula or InstallerStep.Destination or InstallerStep.Diagnostics;
    public bool CanShowFinishButton => CurrentStep == InstallerStep.Completed;

    public string NextButtonText => CurrentStep switch
    {
        InstallerStep.Welcome => "Continue >",
        InstallerStep.Eula => "Continue >",
        InstallerStep.Destination => "Continue >",
        InstallerStep.Diagnostics => "Install Now",
        InstallerStep.Completed => "Finish",
        InstallerStep.Uninstall => "Uninstall Now",
        _ => "Continue >"
    };

    public ICommand NextStepCommand { get; }
    public ICommand PrevStepCommand { get; }
    public ICommand CancelCommand { get; }
    public ICommand BrowsePathCommand { get; }
    public ICommand FinishCommand { get; }
    public ICommand StartUninstallCommand { get; }

    private bool CanGoNext()
    {
        return CurrentStep switch
        {
            InstallerStep.Welcome => true,
            InstallerStep.Eula => IsEulaAccepted,
            InstallerStep.Destination => !string.IsNullOrWhiteSpace(InstallPath),
            InstallerStep.Diagnostics => true,
            _ => false
        };
    }

    private bool CanGoPrev()
    {
        return CurrentStep switch
        {
            InstallerStep.Eula => true,
            InstallerStep.Destination => true,
            InstallerStep.Diagnostics => true,
            _ => false
        };
    }

    public void NextStep()
    {
        if (CurrentStep == InstallerStep.Welcome)
        {
            CurrentStep = InstallerStep.Eula;
        }
        else if (CurrentStep == InstallerStep.Eula)
        {
            CurrentStep = InstallerStep.Destination;
        }
        else if (CurrentStep == InstallerStep.Destination)
        {
            RefreshDiagnostics();
            CurrentStep = InstallerStep.Diagnostics;
        }
        else if (CurrentStep == InstallerStep.Diagnostics)
        {
            _ = RunInstallAsync();
        }
    }

    public void PrevStep()
    {
        if (CurrentStep == InstallerStep.Eula)
        {
            CurrentStep = InstallerStep.Welcome;
        }
        else if (CurrentStep == InstallerStep.Destination)
        {
            CurrentStep = InstallerStep.Eula;
        }
        else if (CurrentStep == InstallerStep.Diagnostics)
        {
            CurrentStep = InstallerStep.Destination;
        }
    }

    #endregion

    #region Configuration Options

    private bool _isEulaAccepted;
    public bool IsEulaAccepted
    {
        get => _isEulaAccepted;
        set
        {
            if (_isEulaAccepted != value)
            {
                _isEulaAccepted = value;
                OnPropertyChanged();
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    private string _installPath = "";
    public string InstallPath
    {
        get => _installPath;
        set
        {
            if (_installPath != value)
            {
                _installPath = value;
                OnPropertyChanged();
            }
        }
    }

    private bool _createDesktopShortcut = true;
    public bool CreateDesktopShortcut
    {
        get => _createDesktopShortcut;
        set { _createDesktopShortcut = value; OnPropertyChanged(); }
    }

    private bool _createStartMenuShortcut = true;
    public bool CreateStartMenuShortcut
    {
        get => _createStartMenuShortcut;
        set { _createStartMenuShortcut = value; OnPropertyChanged(); }
    }

    private bool _launchAppOnFinish = true;
    public bool LaunchAppOnFinish
    {
        get => _launchAppOnFinish;
        set { _launchAppOnFinish = value; OnPropertyChanged(); }
    }

    private void BrowseDestination()
    {
        using var dlg = new System.Windows.Forms.FolderBrowserDialog
        {
            Description = "Select DocLayer installation destination folder",
            SelectedPath = InstallPath,
            UseDescriptionForTitle = true
        };

        if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            InstallPath = Path.Combine(dlg.SelectedPath, "DocLayer");
            RefreshDiagnostics();
        }
    }

    #endregion

    #region Diagnostics

    private List<DiagnosticCheck> _diagnostics = new();
    public List<DiagnosticCheck> Diagnostics
    {
        get => _diagnostics;
        private set { _diagnostics = value; OnPropertyChanged(); }
    }

    public void RefreshDiagnostics()
    {
        Diagnostics = InstallService.RunDiagnostics(InstallPath);
    }

    #endregion

    #region Progress & Execution

    private int _progressPercent;
    public int ProgressPercent
    {
        get => _progressPercent;
        set { _progressPercent = value; OnPropertyChanged(); }
    }

    private string _statusText = "Ready to install...";
    public string StatusText
    {
        get => _statusText;
        set { _statusText = value; OnPropertyChanged(); }
    }

    private bool _hasError;
    public bool HasError
    {
        get => _hasError;
        set { _hasError = value; OnPropertyChanged(); }
    }

    private string _errorMessage = string.Empty;
    public string ErrorMessage
    {
        get => _errorMessage;
        set { _errorMessage = value; OnPropertyChanged(); }
    }

    private async Task RunInstallAsync()
    {
        CurrentStep = InstallerStep.Installing;
        HasError = false;
        ProgressPercent = 0;

        var progress = new Progress<(int Percent, string Status)>(p =>
        {
            ProgressPercent = p.Percent;
            StatusText = p.Status;
        });

        try
        {
            await _installer.InstallAsync(InstallPath, CreateDesktopShortcut, CreateStartMenuShortcut, progress);
            await Task.Delay(400);
            CurrentStep = InstallerStep.Completed;
        }
        catch (Exception ex)
        {
            HasError = true;
            ErrorMessage = "Installation failed: " + ex.Message;
        }
    }

    private async Task RunUninstallAsync()
    {
        ProgressPercent = 0;
        var progress = new Progress<(int Percent, string Status)>(p =>
        {
            ProgressPercent = p.Percent;
            StatusText = p.Status;
        });

        try
        {
            await _installer.UninstallAsync(progress);
            await Task.Delay(400);
            CurrentStep = InstallerStep.Uninstalled;
        }
        catch (Exception ex)
        {
            HasError = true;
            ErrorMessage = "Uninstallation error: " + ex.Message;
        }
    }

    private void Finish()
    {
        if (LaunchAppOnFinish)
        {
            try
            {
                string mainExe = Path.Combine(InstallPath, "Doclayer.exe");
                if (File.Exists(mainExe))
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = mainExe,
                        WorkingDirectory = InstallPath,
                        UseShellExecute = true
                    });
                }
            }
            catch { }
        }

        System.Windows.Application.Current.Shutdown();
    }

    #endregion

    #region EULA Text

    public string EulaText => @"DOCLAYER END-USER LICENSE AGREEMENT (EULA)

Last Updated: 2026

1. GRANT OF LICENSE
DocLayer grants you a personal, non-transferable, and non-exclusive license to use the DocLayer software application on Windows desktop computers running compatible versions of Microsoft Windows and Microsoft Word.

2. MICROSOFT WORD DEPENDENCY & INTEGRATION
DocLayer is designed as a high-performance programmable automation layer for Microsoft Word documents and utilizes Microsoft Word's native COM automation and DISPLAYBARCODE field generation. 
You acknowledge that:
(a) You must possess a valid, properly licensed copy of Microsoft Word to execute COM automation.
(b) DocLayer is an independent product and is not endorsed or sponsored by Microsoft Corporation.

3. RESTRICTIONS
You may not:
- Reverse engineer, decompile, or disassemble the binary packages except as permitted by applicable law.
- Sublicense, lease, or rent the software without written authorization.
- Modify or remove any proprietary notices or branding.

4. DISCLAIMER OF WARRANTIES
THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, AND NON-INFRINGEMENT.

5. LIMITATION OF LIABILITY
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES, OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT, OR OTHERWISE, ARISING FROM, OUT OF, OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

By selecting ""I accept the agreement"" and clicking Next, you confirm your acceptance of these terms.";

    #endregion

    public event PropertyChangedEventHandler? PropertyChanged;

    private void OnPropertyChanged([CallerMemberName] string? n = null)
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
