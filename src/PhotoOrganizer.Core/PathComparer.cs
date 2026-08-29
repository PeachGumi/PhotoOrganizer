namespace PhotoOrganizer.Core;

internal sealed class PathComparer : IComparer<string>, IEqualityComparer<string>
{
    public static PathComparer Instance { get; } = new();

    private static StringComparer Comparer => OperatingSystem.IsWindows()
        ? StringComparer.OrdinalIgnoreCase
        : StringComparer.Ordinal;

    public int Compare(string? x, string? y) => Comparer.Compare(x, y);
    public bool Equals(string? x, string? y) => Comparer.Equals(x, y);
    public int GetHashCode(string obj) => Comparer.GetHashCode(obj);
}
