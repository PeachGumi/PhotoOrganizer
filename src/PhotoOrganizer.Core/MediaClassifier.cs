namespace PhotoOrganizer.Core;

public enum MediaKind
{
    Raw,
    Jpeg,
    Video
}

public sealed class MediaClassifier
{
    private static readonly HashSet<string> StandardJpeg = new(StringComparer.OrdinalIgnoreCase)
    {
        ".jpg", ".jpeg"
    };

    private static readonly HashSet<string> StandardVideo = new(StringComparer.OrdinalIgnoreCase)
    {
        ".mov", ".mp4"
    };

    private readonly HashSet<string> _rawExtensions;

    public MediaClassifier(IEnumerable<string>? rawExtensions = null)
    {
        rawExtensions ??= [".arw", ".cr2", ".cr3", ".nef", ".dng", ".raf", ".rw2", ".orf", ".pef"];
        _rawExtensions = new HashSet<string>(
            rawExtensions.Select(NormalizeExtension),
            StringComparer.OrdinalIgnoreCase);
    }

    public MediaKind? Classify(string path)
    {
        var extension = NormalizeExtension(Path.GetExtension(path));

        // Standard formats are reserved and cannot be reclassified by RAW configuration.
        if (StandardJpeg.Contains(extension)) return MediaKind.Jpeg;
        if (StandardVideo.Contains(extension)) return MediaKind.Video;
        if (_rawExtensions.Contains(extension)) return MediaKind.Raw;
        return null;
    }

    public bool IsSupported(string path) => Classify(path) is not null;

    public static string FolderName(MediaKind kind) => kind switch
    {
        MediaKind.Raw => "RAW",
        MediaKind.Jpeg => "JPG",
        MediaKind.Video => "MP4",
        _ => throw new ArgumentOutOfRangeException(nameof(kind), kind, null)
    };

    private static string NormalizeExtension(string extension)
    {
        if (string.IsNullOrWhiteSpace(extension)) return string.Empty;
        var trimmed = extension.Trim();
        return trimmed.StartsWith('.') ? trimmed.ToLowerInvariant() : $".{trimmed.ToLowerInvariant()}";
    }
}
