namespace PhotoOrganizer.Core;

public enum ImportSafetyStatus
{
    Ready,
    Copying,
    Verifying,
    SafeToReuse,
    Blocked
}

public enum ImportProgressPhase
{
    Scanning,
    Copying,
    Rescanning,
    Verifying
}

public enum ScanFailureReason
{
    InvalidCardRoot,
    MissingMountSessionIdentity,
    MissingPhysicalDeviceIdentity,
    StorageChanged,
    IncompleteScan,
    NoSupportedMedia
}

public sealed record ImportProgress(ImportProgressPhase Phase, int Current, int Total, string Message);

public sealed record ImportScanSession(
    string CardRoot,
    StorageSessionIdentity SourceIdentity,
    IReadOnlyList<string> Files);

public sealed record ScanSessionResult(
    ImportSafetyStatus Status,
    ImportScanSession? Session,
    string Message,
    IReadOnlyList<string> Errors)
{
    public ScanFailureReason? FailureReason { get; init; }
    public bool IsReady => Status == ImportSafetyStatus.Ready && Session is not null;
    public bool IsNoSupportedMedia => FailureReason == ScanFailureReason.NoSupportedMedia;
}

public sealed record ImportSummary(
    int TotalSupported,
    int Copied,
    int SkippedAlreadyBackedUp,
    int Failed,
    string BasePath,
    IReadOnlyList<string> Warnings,
    IReadOnlyList<string> Errors);

public sealed record ImportRunResult(
    ImportSafetyStatus Status,
    string Message,
    ImportSummary Summary,
    FormatVerificationResult? Verification = null)
{
    public bool IsSafeToReuse => Status == ImportSafetyStatus.SafeToReuse;
}
