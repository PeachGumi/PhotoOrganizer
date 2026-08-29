namespace PhotoOrganizer.Core;

public static class PathSafety
{
    public static string Normalize(string path)
    {
        if (string.IsNullOrWhiteSpace(path)) return string.Empty;

        var full = Path.GetFullPath(path);
        var root = Path.GetPathRoot(full);
        if (!string.IsNullOrEmpty(root) && string.Equals(full, root, StringComparison.OrdinalIgnoreCase))
        {
            return root;
        }

        return full.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
    }

    public static bool IsSameOrDescendant(string path, string ancestor, StringComparison comparison)
    {
        var normalizedPath = Normalize(path);
        var normalizedAncestor = Normalize(ancestor);
        if (string.Equals(normalizedPath, normalizedAncestor, comparison)) return true;

        var prefix = normalizedAncestor.EndsWith(Path.DirectorySeparatorChar)
            || normalizedAncestor.EndsWith(Path.AltDirectorySeparatorChar)
            ? normalizedAncestor
            : normalizedAncestor + Path.DirectorySeparatorChar;

        return normalizedPath.StartsWith(prefix, comparison);
    }

    /// <summary>
    /// Verifies that an absolute path is reached only through direct filesystem
    /// components. A symlink/junction/reparse point can make a lexically separate
    /// destination resolve back onto the camera card, so user-controlled aliases are
    /// never permitted to participate in backup or reuse-safety decisions.
    ///
    /// Missing leaf components are allowed because the application may create them;
    /// every component that already exists is inspected, including its ancestors.
    /// Any metadata error other than a genuinely missing component fails closed.
    /// </summary>
    public static bool TryValidateDirectFilesystemPath(string path, out string? error)
    {
        error = null;

        string current;
        try
        {
            current = Normalize(path);
        }
        catch (Exception ex)
        {
            error = $"Path could not be normalized: {ex.Message}";
            return false;
        }

        if (string.IsNullOrWhiteSpace(current))
        {
            error = "Path is empty.";
            return false;
        }

        while (true)
        {
            try
            {
                var attributes = File.GetAttributes(current);
                if ((attributes & FileAttributes.ReparsePoint) != 0
                    && !IsVerifiedMacSystemAlias(current))
                {
                    error = $"Path contains a symbolic link, junction, or reparse point: {current}";
                    return false;
                }
            }
            catch (FileNotFoundException)
            {
                // A not-yet-created leaf is allowed; existing ancestors are still checked below.
            }
            catch (DirectoryNotFoundException)
            {
                // A not-yet-created leaf is allowed; existing ancestors are still checked below.
            }
            catch (Exception ex)
            {
                error = $"Unable to inspect path component {current}: {ex.Message}";
                return false;
            }

            string? parent;
            try
            {
                parent = Directory.GetParent(current)?.FullName;
            }
            catch (Exception ex)
            {
                error = $"Unable to inspect path parent for {current}: {ex.Message}";
                return false;
            }

            if (string.IsNullOrWhiteSpace(parent)
                || string.Equals(parent, current, OperatingSystem.IsWindows()
                    ? StringComparison.OrdinalIgnoreCase
                    : StringComparison.Ordinal))
            {
                break;
            }

            current = parent;
        }

        return true;
    }

    private static bool IsVerifiedMacSystemAlias(string path)
    {
        // macOS exposes /var as a root-owned compatibility symlink to /private/var.
        // .NET temporary directories therefore commonly live under /var/folders.
        // Trust only this exact OS alias and only after resolving it to the expected
        // fixed target; every user-controlled alias remains fail-closed.
        if (!OperatingSystem.IsMacOS() || !string.Equals(path, "/var", StringComparison.Ordinal))
        {
            return false;
        }

        try
        {
            var target = new DirectoryInfo(path).ResolveLinkTarget(returnFinalTarget: true);
            return target is not null
                && string.Equals(Normalize(target.FullName), "/private/var", StringComparison.Ordinal);
        }
        catch
        {
            return false;
        }
    }
}
