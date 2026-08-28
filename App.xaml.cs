using System.IO;
using System.Windows;
using System.Windows.Media;

namespace WordBarcodeStudio;

public partial class App : System.Windows.Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        ThemeManager.Initialize();
        LoadLogoFromSvg();

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

    /// <summary>
    /// Loads Assets/logo.svg at runtime (via SharpVectors) and replaces the fallback
    /// LogoDrawing resource, so edits to the SVG file are reflected in the app
    /// after a rebuild.
    /// </summary>
    private static void LoadLogoFromSvg()
    {
        try
        {
            var svgPath = Path.Combine(AppContext.BaseDirectory, "Assets", "logo.svg");
            if (!File.Exists(svgPath)) return;

            var settings = new SharpVectors.Renderers.Wpf.WpfDrawingSettings();
            var converter = new SharpVectors.Converters.FileSvgConverter(settings);
            if (converter.Convert(svgPath) && converter.Drawing != null)
            {
                var drawingImage = new DrawingImage(converter.Drawing);
                drawingImage.Freeze();
                System.Windows.Application.Current.Resources["LogoDrawing"] = drawingImage;
            }
        }
        catch
        {
            // keep the fallback LogoDrawing from App.xaml
        }
    }
}
