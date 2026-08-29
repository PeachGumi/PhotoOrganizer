using PhotoOrganizer.Core;

namespace PhotoOrganizer.App;

public static class StartupCardDiscovery
{
    public static async Task InitializeAsync(
        MainWindowViewModel viewModel,
        CameraCardRootResolver cardRoots)
    {
        ArgumentNullException.ThrowIfNull(viewModel);
        ArgumentNullException.ThrowIfNull(cardRoots);

        IReadOnlyList<string> candidates;
        try
        {
            // Platform enumeration may invoke diskutil/WMI. Keep it away from the
            // Avalonia dispatcher just like the individual card scans below.
            candidates = await Task.Run(cardRoots.GetCandidateRoots).ConfigureAwait(true);
        }
        catch (Exception exception)
        {
            viewModel.ReportUiFailure("起動時のSDカード検出", exception);
            return;
        }

        foreach (var candidate in candidates)
        {
            if (viewModel.IsBusy || !string.IsNullOrWhiteSpace(viewModel.SelectedSdPath))
            {
                return;
            }

            await viewModel.ScanCardAsync(candidate, autoDetected: true).ConfigureAwait(true);

            if (!string.IsNullOrWhiteSpace(viewModel.SelectedSdPath))
            {
                return;
            }

            // Auto-detected cards with no supported media deliberately return to the
            // waiting state. Only that benign outcome should advance to the next card;
            // cancellation or a real scan failure remains visible and fail-closed.
            if (!string.Equals(viewModel.ProgressLabel, "待機中", StringComparison.Ordinal))
            {
                return;
            }
        }
    }
}
