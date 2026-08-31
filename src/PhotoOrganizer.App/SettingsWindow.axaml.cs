using Avalonia.Controls;
using Avalonia.Interactivity;

namespace PhotoOrganizer.App;

public sealed partial class SettingsWindow : Window
{
    private readonly MainWindowViewModel? _viewModel;
    private readonly BackgroundPreferencesStore _backgroundPreferences = new();
    private readonly IStartupRegistrationService _startupRegistration = new StartupRegistrationService();
    private bool _changingBackgroundPreference;

    public SettingsWindow()
    {
        InitializeComponent();
    }

    public SettingsWindow(MainWindowViewModel viewModel) : this()
    {
        _viewModel = viewModel;
        DataContext = viewModel;
        StartInBackgroundCheckBox.Click += StartInBackground_Changed;
        Closed += OnClosed;
        LoadBackgroundPreference();
    }

    private void OnClosed(object? sender, EventArgs e)
    {
        StartInBackgroundCheckBox.Click -= StartInBackground_Changed;
    }

    private void StartInBackground_Changed(object? sender, RoutedEventArgs e)
    {
        if (_changingBackgroundPreference || _viewModel is null) return;

        var desired = StartInBackgroundCheckBox.IsChecked == true;
        var previous = _backgroundPreferences.Load().StartInBackground;
        if (desired == previous)
        {
            UpdateBackgroundSettingStatus(desired, null);
            return;
        }

        if (!_backgroundPreferences.Save(desired, out var error))
        {
            RevertBackgroundCheckBox(previous);
            UpdateBackgroundSettingStatus(previous, $"設定保存失敗: {error}");
            return;
        }

        if (_viewModel.AutoStart && !_startupRegistration.SetEnabled(true, out var startupError))
        {
            _backgroundPreferences.Save(previous, out _);
            RevertBackgroundCheckBox(previous);
            UpdateBackgroundSettingStatus(previous, $"自動起動設定の更新失敗: {startupError}");
            return;
        }

        UpdateBackgroundSettingStatus(desired, null);
    }

    private void LoadBackgroundPreference()
    {
        var enabled = _backgroundPreferences.Load().StartInBackground;
        RevertBackgroundCheckBox(enabled);
        UpdateBackgroundSettingStatus(enabled, null);
    }

    private void RevertBackgroundCheckBox(bool value)
    {
        _changingBackgroundPreference = true;
        StartInBackgroundCheckBox.IsChecked = value;
        _changingBackgroundPreference = false;
    }

    private void UpdateBackgroundSettingStatus(bool enabled, string? error)
    {
        BackgroundSettingStatus.Text = error ?? (enabled
            ? "次回はメニューバー常駐で起動し、SDカード検出時に画面を表示します。"
            : "次回起動時にメイン画面を表示します。");
    }
}
