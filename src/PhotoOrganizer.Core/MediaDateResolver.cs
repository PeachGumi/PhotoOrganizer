using MetadataExtractor;
using MetadataExtractor.Formats.Exif;

namespace PhotoOrganizer.Core;

public sealed class MediaDateResolver
{
    public DateTime ResolveDate(string path)
    {
        var exif = TryReadExifDate(path);
        if (exif is not null) return exif.Value;

        try
        {
            var lastWrite = File.GetLastWriteTime(path);
            if (lastWrite > DateTime.MinValue) return lastWrite;
        }
        catch
        {
            // Fall through to a deterministic current-date fallback.
        }

        return DateTime.Now;
    }

    public string ResolveDateKey(IEnumerable<string> files)
    {
        var dates = files
            .Select(ResolveDate)
            .OrderBy(date => date)
            .ToArray();

        return (dates.Length == 0 ? DateTime.Now : dates[0]).ToString("yyyy-MM-dd");
    }

    public IReadOnlyList<string> ResolveDateKeys(IEnumerable<string> files)
    {
        return files
            .Select(path => ResolveDate(path).ToString("yyyy-MM-dd"))
            .Distinct(StringComparer.Ordinal)
            .OrderBy(key => key, StringComparer.Ordinal)
            .ToArray();
    }

    private static DateTime? TryReadExifDate(string path)
    {
        try
        {
            var directories = ImageMetadataReader.ReadMetadata(path);
            var exif = directories.OfType<ExifSubIfdDirectory>().FirstOrDefault();
            if (exif is null) return null;

            if (exif.TryGetDateTime(ExifDirectoryBase.TagDateTimeOriginal, out var original))
            {
                return original;
            }

            if (exif.TryGetDateTime(ExifDirectoryBase.TagDateTimeDigitized, out var digitized))
            {
                return digitized;
            }
        }
        catch
        {
            // RAW/video formats unsupported by the metadata library use filesystem date fallback.
        }

        return null;
    }
}
