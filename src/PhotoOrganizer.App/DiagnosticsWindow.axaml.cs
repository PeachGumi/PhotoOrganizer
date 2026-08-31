using Avalonia.Controls;

namespace PhotoOrganizer.App;

public sealed partial class DiagnosticsWindow : Window
{
    public DiagnosticsWindow()
    {
        InitializeComponent();
    }

    public DiagnosticsWindow(MainWindowViewModel viewModel) : this()
    {
        DataContext = viewModel;
    }
}
