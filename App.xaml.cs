using System.Windows;

namespace WordBarcodeStudio;

public partial class App : System.Windows.Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        ThemeManager.Initialize();

        DispatcherUnhandledException += (_, ev) =>
        {
            System.Windows.MessageBox.Show(
                ev.Exception.ToString(),
                "Doclayer - Unhandled Error",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            ev.Handled = true;
        };
    }
}
