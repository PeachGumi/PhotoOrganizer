namespace PhotoOrganizer.Core;

public sealed class ImportCoordinator
{
    private static readonly int MaxConcurrentIoOperations = Math.Clamp(
        Environment.ProcessorCount,
        1,
        4);

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
        CameraCardRootResolver cardRoots,
        IStorageVolumeProvider volumeProvider)
    {
        _classifier = classifier;
        _scanner = new MediaScanner(classifier, volumeProvider);
        _copyService = new SafeCopyService();
        _verifier = new FormatSafetyVerifier(
            classifier,
            volumeProvider,
            maxDegreeOfParallelism: MaxConcurrentIoOperations);
        _destinationLibrary = new DestinationLibrary(
            volumeProvider,
            maxDegreeOfParallelism: MaxConcurrentIoOperations);
        _dateResolver = new MediaDateResolver();
        _storageSessions = storageSessions;
        _cardRoots = cardRoots;
    }

    public ScanSessionResult ScanCard(
        string selectedPath,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var root = _cardRoots.Resolve(selectedPath);
        if (root is null)
        {
            return BlockedScan(
                "A complete camera-card root containing DCIM or PRIVATE could not be established.",
                ScanFailureReason.InvalidCardRoot);
        }

        var identity = _storageSessions.Capture(root);
        if (identity is null)
        {
            return BlockedScan(
                "The camera-card mount-session identity is unavailable.",
                ScanFailureReason.MissingMountSessionIdentity);
        }

        if (string.IsNullOrWhiteSpace(identity.PhysicalDeviceFingerprint))
        {
            return BlockedScan(
                "The camera-card physical-device identity is unavailable.",
                ScanFailureReason.MissingPhysicalDeviceIdentity);
        }

        var scan = _scanner.Scan(root, cancellationToken);
        if (!_storageSessions.Matches(identity, root))
        {
            return BlockedScan(
                "The camera-card volume changed while it was being scanned.",
                ScanFailureReason.StorageChanged);
        }

        if (!scan.IsComplete)
        {
            return BlockedScan(
                "The camera card could not be scanned completely.",
                ScanFailureReason.IncompleteScan,
                scan.Errors);
        }

        if (scan.Files.Count == 0)
        {
            return BlockedScan(
                "No supported media exists on the selected camera card.",
                ScanFailureReason.NoSupportedMedia);
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
        var summary = new ImportSummary(session.Files.Count, 0, 0, session.Files.Count, string.Empty, warnings, errors);

        try
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (string.IsNullOrWhiteSpace(session.SourceIdentity.PhysicalDeviceFingerprint))
            {
                return Blocked("The camera-card physical-device identity is unavailable. Scan it again before importing.", summary);
            }

            if (!_storageSessions.Matches(session.SourceIdentity, session.CardRoot))
            {
                return Blocked("The camera card changed after the successful scan. Scan it again.", summary);
            }

            if (string.IsNullOrWhiteSpace(destinationRoot))
            {
                return Blocked("Destination is empty.", summary);
            }

            if (string.IsNullOrWhiteSpace(eventName))
            {
                return Blocked("Event name is empty.", summary);
            }

            var destination = PathSafety.Normalize(destinationRoot);
            if (PathSafety.IsSameOrDescendant(destination, session.CardRoot, _storageSessions.PathComparison)
                || PathSafety.IsSameOrDescendant(session.CardRoot, destination, _storageSessions.PathComparison))
            {
                return Blocked("Destination must not be the camera card or its parent/child path.", summary);
            }

            var destinationIdentity = _storageSessions.Capture(destination);
            if (destinationIdentity is null)
            {
                return Blocked("Destination mount-session identity is unavailable.", summary);
            }

            if (string.IsNullOrWhiteSpace(destinationIdentity.PhysicalDeviceFingerprint))
            {
                return Blocked("Destination physical-device identity is unavailable. Choose a destination whose physical storage can be verified.", summary);
            }

            if (destinationIdentity.SessionId == session.SourceIdentity.SessionId
                || string.Equals(destinationIdentity.Fingerprint, session.SourceIdentity.Fingerprint, StringComparison.Ordinal))
            {
                return Blocked("Destination must be on a different volume from the camera card.", summary);
            }

            if (string.Equals(
                    destinationIdentity.PhysicalDeviceFingerprint,
                    session.SourceIdentity.PhysicalDeviceFingerprint,
                    StringComparison.Ordinal))
            {
                return Blocked("Destination must be on a different physical storage device from the camera card.", summary);
            }

            var initialFiles = session.Files
                .Where(_classifier.IsSupported)
                .Select(Path.GetFullPath)
                .ToHashSet(PathComparer.Instance);

            if (initialFiles.Count == 0)
            {
                return Blocked("The scan session contains no supported media.", summary);
            }

            foreach (var source in initialFiles)
            {
                if (!PathSafety.IsSameOrDescendant(source, session.CardRoot, _storageSessions.PathComparison)
                    || !PathSafety.TryValidateDirectFilesystemPath(source, out _))
                {
                    return Blocked(
                        "The scan session contains media outside the direct camera-card filesystem path. Scan the card again.",
                        summary);
                }
            }

            var lookup = await _destinationLibrary
                .FindVerifiedBackupsAsync(initialFiles, destination, cancellationToken)
                .ConfigureAwait(false);

            if (lookup.Errors.Count > 0)
            {
                warnings.Add($"Destination library lookup reported {lookup.Errors.Count} error(s). Final reuse verification remains fail-closed.");
            }

            var pending = initialFiles
                .Where(path => !lookup.MatchedSources.Contains(path))
                .OrderBy(path => path, PathComparer.Instance)
                .ToArray();

            var skipped = initialFiles.Count - pending.Length;
            if (skipped > 0)
            {
                warnings.Add($"Skipped {skipped} file(s) already proven byte-identical in the destination library.");
            }

            IEnumerable<string> dateFiles = pending.Length == 0 ? initialFiles : pending;
            var dateResolution = _dateResolver.ResolveDateSummary(dateFiles);
            var dateKey = dateResolution.EarliestDateKey;
            var dateKeys = dateResolution.DateKeys;
            if (dateKeys.Count > 1)
            {
                warnings.Add($"Multiple capture dates detected ({string.Join(", ", dateKeys)}); using earliest date {dateKey}.");
            }

            var sanitizedEvent = SanitizeEventName(eventName);
            if (string.IsNullOrWhiteSpace(sanitizedEvent))
            {
                return Blocked("Event name contains no usable filename characters.", summary);
            }

            var basePath = Path.Combine(destination, dateKey[..4], $"{dateKey}_{sanitizedEvent}");
            summary = summary with
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
            void ApplyCopyResults(IEnumerable<PendingCopyResult> completed)
            {
                foreach (var completedCopy in completed)
                {
                    if (completedCopy.Result.Status == CopyStatus.Failed)
                    {
                        failed++;
                        errors.Add($"{completedCopy.Source}: {completedCopy.Result.Error ?? "copy failed"}");
                    }
                    else if (completedCopy.Result.Status == CopyStatus.Copied)
                    {
                        copied++;
                    }
                    else
                    {
                        skipped++;
                    }
                }
            }

            for (var batchStart = 0; batchStart < pending.Length; batchStart += MaxConcurrentIoOperations)
            {
                var batchEnd = Math.Min(batchStart + MaxConcurrentIoOperations, pending.Length);
                var batch = new List<Task<PendingCopyResult>>(batchEnd - batchStart);

                progress?.Report(new ImportProgress(
                    ImportProgressPhase.Copying,
                    batchStart,
                    pending.Length,
                    $"Copying {batchStart}/{pending.Length}"));

                for (var index = batchStart; index < batchEnd; index++)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    if (!StorageMatches(session, destinationIdentity, destination))
                    {
                        ApplyCopyResults(await Task.WhenAll(batch).ConfigureAwait(false));
                        return Blocked(
                            "Source or destination storage changed during copying.",
                            summary with
                            {
                                Copied = copied,
                                SkippedAlreadyBackedUp = skipped,
                                Failed = failed + pending.Length - index
                            });
                    }

                    var source = pending[index];
                    var kind = _classifier.Classify(source);
                    if (kind is null)
                    {
                        failed++;
                        errors.Add($"{source}: file became unsupported unexpectedly.");
                        continue;
                    }

                    batch.Add(CopyPendingAsync(
                        source,
                        Path.Combine(basePath, MediaClassifier.FolderName(kind.Value)),
                        cancellationToken));
                }

                ApplyCopyResults(await Task.WhenAll(batch).ConfigureAwait(false));
                progress?.Report(new ImportProgress(
                    ImportProgressPhase.Copying,
                    batchEnd,
                    pending.Length,
                    $"Copying {batchEnd}/{pending.Length}"));
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
                "Copy processing complete; camera-card reuse is not approved yet."));

            if (failed > 0)
            {
                return Blocked("One or more supported media files failed to copy. Do not reuse the card.", summary);
            }

            if (!StorageMatches(session, destinationIdentity, destination))
            {
                return Blocked("Source or destination storage changed after copying.", summary);
            }

            progress?.Report(new ImportProgress(
                ImportProgressPhase.Rescanning,
                0,
                initialFiles.Count,
                "Rescanning the complete camera card."));

            var rescan = _scanner.Scan(session.CardRoot, cancellationToken);

            if (!StorageMatches(session, destinationIdentity, destination))
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
                "Verifying destination SHA-256 and durable storage commit."));

            var verification = await _verifier
                .VerifyAsync(rescan.Files, destination, cancellationToken)
                .ConfigureAwait(false);

            if (!StorageMatches(session, destinationIdentity, destination))
            {
                return Blocked("Source or destination storage changed during final byte and durability verification.", summary, verification);
            }

            if (!verification.IsSafe)
            {
                errors.AddRange(verification.Errors);
                return Blocked(
                    $"Only {verification.Verified} of {verification.Total} supported file(s) were verified and durably synchronized in the destination. Do not reuse the card.",
                    summary,
                    verification);
            }

            // Verification can take long enough for another supported file to appear
            // after the post-copy scan. A final whole-card enumeration closes that
            // window without repeating the expensive SHA-256/durability pass.
            var finalRescan = _scanner.Scan(session.CardRoot, cancellationToken);

            if (!StorageMatches(session, destinationIdentity, destination))
            {
                return Blocked(
                    "Source or destination storage changed during the final camera-card consistency scan.",
                    summary,
                    verification);
            }

            if (!finalRescan.IsComplete)
            {
                errors.AddRange(finalRescan.Errors);
                return Blocked(
                    "Final camera-card consistency scan was incomplete. Do not reuse the card.",
                    summary,
                    verification);
            }

            var finalFiles = finalRescan.Files
                .Select(Path.GetFullPath)
                .ToHashSet(PathComparer.Instance);
            if (!rescanned.SetEquals(finalFiles))
            {
                return Blocked(
                    "Supported media changed during final verification. Do not reuse the card.",
                    summary,
                    verification);
            }

            progress?.Report(new ImportProgress(
                ImportProgressPhase.Verifying,
                verification.Verified,
                verification.Total,
                "Destination copies verified and durably synchronized; camera card may be reused."));

            return new ImportRunResult(
                ImportSafetyStatus.SafeToReuse,
                $"Verified {verification.Verified} supported file(s) by size and SHA-256 and durably synchronized destination storage. Camera card may be reused.",
                summary,
                verification);
        }
        catch (OperationCanceledException)
        {
            return Blocked("Import was cancelled before final verification. Do not reuse the card.", summary);
        }
        catch (Exception ex)
        {
            errors.Add(ex.Message);
            return Blocked("Import failed before final verification. Do not reuse the card.", summary);
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

    private bool StorageMatches(
        ImportScanSession session,
        StorageSessionIdentity destinationIdentity,
        string destination) =>
        _storageSessions.MatchesPair(
            session.SourceIdentity,
            session.CardRoot,
            destinationIdentity,
            destination);

    private async Task<PendingCopyResult> CopyPendingAsync(
        string source,
        string destinationDirectory,
        CancellationToken cancellationToken)
    {
        var result = await _copyService
            .CopyAsync(source, destinationDirectory, cancellationToken)
            .ConfigureAwait(false);
        return new PendingCopyResult(source, result);
    }

    private sealed record PendingCopyResult(string Source, CopyResult Result);

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

    private static ScanSessionResult BlockedScan(
        string message,
        ScanFailureReason failureReason,
        IReadOnlyList<string>? errors = null) =>
        new(ImportSafetyStatus.Blocked, null, message, errors ?? [])
        {
            FailureReason = failureReason
        };

    private static ImportRunResult Blocked(
        string message,
        ImportSummary summary,
        FormatVerificationResult? verification = null) =>
        new(ImportSafetyStatus.Blocked, message, summary, verification);
}
