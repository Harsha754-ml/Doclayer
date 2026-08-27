using System.Windows;
using WordBarcodeStudio.ViewModels;

namespace WordBarcodeStudio.Views;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
        DataContext = new MainViewModel();
        Closing += (_, _) => ((MainViewModel)DataContext).Dispose();
    }
}
