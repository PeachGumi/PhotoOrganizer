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

public sealed record ImportProgress(
    ImportProgressPhase Phase,
    int Current,
    int Total,
    string Message);

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
    public bool IsReady => Status == ImportSafetyStatus.Ready && Session is not null;
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

public sealed class ImportCoordinator
{
    private readonly MediaClassifier _classifier;
    private readonly MediaScanner _scanner;
    private readonly SafeCopyService _copyService;
    private readonly FormatSafetyVerifier _verifier;
    private readonly DestinationLibrary _destinationLibrary;
    private readonly MediaDateResolver _dateResolver;
    private readonly StorageSessionTracker _storageSessions;
    private readonly CameraCardRootResolver _cardRoots;

    public ImportCoordinator(
        MediaClassifier classifier,
        StorageSessionTracker storageSessions,
        CameraCardRootResolver cardRoots)
    {
        _classifier = classifier;
        _scanner = new MediaScanner(classifier);
        _copyService = new SafeCopyService();
        _verifier = new FormatSafetyVerifier(classifier);
        _destinationLibrary = new DestinationLibrary();
        _dateResolver = new MediaDateResolver();
        _storageSessions = storageSessions;
        _cardRoots = cardRoots;
    }

    public ScanSessionResult ScanCard(string selectedPath)
    {
        var root = _cardRoots.Resolve(selectedPath);
        if (root is null)
        {
            return BlockedScan("A complete camera-card root containing DCIM or PRIVATE could not be established.");
        }

        var identity = _storageSessions.Capture(root);
        if (identity is null)
        {
            return BlockedScan("The camera-card mount-session identity is unavailable.");
        }

        var scan = _scanner.Scan(root);
        if (!_storageSessions.Matches(identity, root))
        {
            return BlockedScan("The camera-card volume changed while it was being scanned.");
        }

        if (!scan.IsComplete)
        {
            return new ScanSessionResult(
                ImportSafetyStatus.Blocked,
                null,
                "The camera card could not be scanned completely.",
                scan.Errors);
        }

        if (scan.Files.Count == 0)
        {
            return BlockedScan("No supported media exists on the selected camera card.");
        }

        return new ScanSessionResult(
            ImportSafetyStatus.Ready,
            new ImportScanSession(root, identity, scan.Files),
            $"Complete scan: {scan.Files.Count} supported file(s).",
            []);
    }

    public async Task<ImportRunResult> ImportAsync(
        ImportScanSession session,
        string destinationRoot,
        string eventName,
        IProgress<ImportProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var warnings = new List<string>();
        var errors = new List<string>();
        var emptySummary = new ImportSummary(session.Files.Count, 0, 0, session.Files.Count, string.Empty, warnings, errors);

        try
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (!_storageSessions.Matches(session.SourceIdentity, session.CardRoot))
            {
                return Blocked("The camera card changed after the successful scan. Scan it again.", emptySummary);
            }

            if (string.IsNullOrWhiteSpace(destinationRoot))
            {
                return Blocked("Destination is empty.", emptySummary);
            }

            if (string.IsNullOrWhiteSpace(eventName))
            {
                return Blocked("Event name is empty.", emptySummary);
            }

            var destination = PathSafety.Normalize(destinationRoot);
            if (PathSafety.IsSameOrDescendant(destination, session.CardRoot, _storageSessions.PathComparison)
                || PathSafety.IsSameOrDescendant(session.CardRoot, destination, _storageSessions.PathComparison))
            {
                return Blocked("Destination must not be the camera card or its parent/child path.", emptySummary);
            }

            var destinationIdentity = _storageSessions.Capture(destination);
            if (destinationIdentity is null)
            {
                return Blocked("Destination mount-session identity is unavailable.", emptySummary);
            }

            if (destinationIdentity.SessionId == session.SourceIdentity.SessionId)
            {
                return Blocked("Destination must be on a different volume from the camera card.", emptySummary);
            }

            var initialFiles = session.Files
                .Where(_classifier.IsSupported)
                .Select(Path.GetFullPath)
                .ToHashSet(PathComparer.Instance);

            var lookup = await _destinationLibrary
                .FindVerifiedBackupsAsync(initialFiles, destination, cancellationToken)
                .ConfigureAwait(false);

            if (lookup.Errors.Count > 0)
            {
                warnings.Add($"Destination library lookup reported {lookup.Errors.Count} error(s); final reuse verification will remain fail-closed.");
            }

            var pending = initialFiles
                .Where(path => !lookup.MatchedSources.Contains(path))
                .OrderBy(path => path, PathComparer.Instance)
                .ToArray();

            var skipped = initialFiles.Count - pending.Length;
            if (skipped > 0)
            {
                warnings.Add($"Skipped {skipped} file(s) already proven byte-identical elsewhere in the destination library.");
            }

            var dateFiles = pending.Length == 0 ? initialFiles : pending;
            var dateKey = _dateResolver.ResolveDateKey(dateFiles);
            var dateKeys = _dateResolver.ResolveDateKeys(dateFiles);
            if (dateKeys.Count > 1)
            {
                warnings.Add($"Multiple capture dates detected ({string.Join(", ", dateKeys)}); using earliest date {dateKey}.");
            }

            var sanitizedEvent = SanitizeEventName(eventName);
            if (string.IsNullOrWhiteSpace(sanitizedEvent))
            {
                return Blocked("Event name contains no usable filename characters.", emptySummary);
            }

            var basePath = Path.Combine(destination, dateKey[..4], $"{dateKey}_{sanitizedEvent}");
            var summary = emptySummary with
            {
                Failed = 0,
                SkippedAlreadyBackedUp = skipped,
                BasePath = basePath
            };

            if (pending.Length > 0)
            {
                var requiredBytes = pending.Sum(path => new FileInfo(path).Length);
                var availableBytes = TryGetAvailableBytes(destinationIdentity.RootPath);
                if (availableBytes is not null && availableBytes.Value < requiredBytes)
                {
                    return Blocked(
                        $"Destination free space is insufficient. Required {requiredBytes} bytes; available {availableBytes.Value} bytes.",
                        summary with { Failed = pending.Length });
                }

                foreach (var kind in Enum.GetValues<MediaKind>())
                {
                    Directory.CreateDirectory(Path.Combine(basePath, MediaClassifier.FolderName(kind)));
                }
            }

            var copied = 0;
            var failed = 0;
            for (var index = 0; index < pending.Length; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!_storageSessions.Matches(session.SourceIdentity, session.CardRoot)
                    || !_storageSessions.Matches(destinationIdentity, destination))
                {
                    return Blocked(
                        "Source or destination storage changed during copying.",
                        summary with { Copied = copied, Failed = pending.Length - index });
                }

                var source = pending[index];
                var kind = _classifier.Classify(source);
                if (kind is null)
                {
                    failed++;
                    errors.Add($"{source}: file became unsupported unexpectedly.");
                    continue;
                }

                progress?.Report(new ImportProgress(
                    ImportProgressPhase.Copying,
                    index,
                    pending.Length,
                    $"Copying {index}/{pending.Length}"));

                var result = await _copyService.CopyAsync(
                    source,
                    Path.Combine(basePath, MediaClassifier.FolderName(kind.Value)),
                    cancellationToken).ConfigureAwait(false);

                if (result.Status == CopyStatus.Failed)
                {
                    failed++;
                    errors.Add($"{source}: {result.Error ?? "copy failed"}");
                }
                else if (result.Status == CopyStatus.Copied)
                {
                    copied++;
                }
                else if (result.Status == CopyStatus.SkippedDuplicate)
                {
                    skipped++;
                }
            }

            summary = summary with
            {
                Copied = copied,
                SkippedAlreadyBackedUp = skipped,
                Failed = failed
            };

            progress?.Report(new ImportProgress(
                ImportProgressPhase.Copying,
                pending.Length,
                pending.Length,
                "Copy processing complete; reuse is not approved yet."));

            if (failed > 0)
            {
                return Blocked("One or more supported media files failed to copy. Do not reuse the card.", summary);
            }

            if (!_storageSessions.Matches(session.SourceIdentity, session.CardRoot)
                || !_storageSessions.Matches(destinationIdentity, destination))
            {
                return Blocked("Source or destination storage changed after copying.", summary);
            }

            progress?.Report(new ImportProgress(
                ImportProgressPhase.Rescanning,
                0,
                initialFiles.Count,
                "Rescanning the complete camera card."));

            var rescan = _scanner.Scan(session.CardRoot);
            if (!_storageSessions.Matches(session.SourceIdentity, session.CardRoot)
                || !_storageSessions.Matches(destinationIdentity, destination))
            {
                return Blocked("Source or destination storage changed during the post-import rescan.", summary);
            }

            if (!rescan.IsComplete)
            {
                errors.AddRange(rescan.Errors);
                return Blocked("Post-import camera-card scan was incomplete. Do not reuse the card.", summary);
            }

            if (rescan.Files.Count == 0)
            {
                return Blocked("No supported media was visible during the post-import rescan.", summary);
            }

            var rescanned = rescan.Files.Select(Path.GetFullPath).ToHashSet(PathComparer.Instance);
            if (!initialFiles.IsSubsetOf(rescanned))
            {
                return Blocked("Supported media visible before import is missing from the post-import rescan.", summary);
            }

            progress?.Report(new ImportProgress(
                ImportProgressPhase.Verifying,
                0,
                rescan.Files.Count,
                "Verifying destination bytes with SHA-256."));

            var verification = await _verifier
                .VerifyAsync(rescan.Files, destination, cancellationToken)
                .ConfigureAwait(false);

            if (!_storageSessions.Matches(session.SourceIdentity, session.CardRoot)
                || !_storageSessions.Matches(destinationIdentity, destination))
            {
                return Blocked("Source or destination storage changed during final byte verification.", summary, verification);
            }

            if (!verification.IsSafe)
            {
                errors.AddRange(verification.Errors);
                return Blocked(
                    $"Only {verification.Verified} of {verification.Total} supported file(s) were verified in the destination. Do not reuse the card.",
                    summary,
                    verification);
            }

            progress?.Report(new ImportProgress(
                ImportProgressPhase.Verifying,
                verification.Verified,
                verification.Total,
                "Destination copies verified; camera card may be reused."));

            return new ImportRunResult(
                ImportSafetyStatus.SafeToReuse,
                $"Verified {verification.Verified} supported file(s) by size and SHA-256. Camera card may be reused.",
                summary,
                verification);
        }
        catch (OperationCanceledException)
        {
            return Blocked("Import was cancelled before final verification. Do not reuse the card.", emptySummary);
        }
        catch (Exception ex)
        {
            errors.Add(ex.Message);
            return Blocked("Import failed before final verification. Do not reuse the card.", emptySummary);
        }
    }

    public static string SanitizeEventName(string name)
    {
        var invalid = Path.GetInvalidFileNameChars()
            .Concat(['/', '\\', ':', '*', '?', '"', '<', '>', '|'])
            .ToHashSet();
        var sanitized = new string(name.Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray());
        return sanitized.Trim();
    }

    private static long? TryGetAvailableBytes(string volumeRoot)
    {
        try
        {
            return new DriveInfo(volumeRoot).AvailableFreeSpace;
        }
        catch
        {
            return null;
        }
    }

    private static ScanSessionResult BlockedScan(string message) =>
        new(ImportSafetyStatus.Blocked, null, message, []);

    private static ImportRunResult Blocked(
        string message,
        ImportSummary summary,
        FormatVerificationResult? verification = null) =>
        new(ImportSafetyStatus.Blocked, message, summary, verification);
}
