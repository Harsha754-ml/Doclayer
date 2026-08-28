using System;
using System.Linq;
using System.Windows;
using DocLayer.Setup.ViewModels;
using DocLayer.Setup.Views;

namespace DocLayer.Setup;

public partial class App : System.Windows.Application
{
    private void Application_Startup(object sender, StartupEventArgs e)
    {
        bool isUninstall = e.Args.Any(a => a.Equals("--uninstall", StringComparison.OrdinalIgnoreCase) ||
                                           a.Equals("/uninstall", StringComparison.OrdinalIgnoreCase) ||
                                           a.Equals("-u", StringComparison.OrdinalIgnoreCase));

        var vm = new InstallerViewModel(isUninstall);
        var window = new MainWindow { DataContext = vm };
        window.Show();
    }
}
