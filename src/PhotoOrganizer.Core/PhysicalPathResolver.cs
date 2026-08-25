namespace PhotoOrganizer.Core;

/// <summary>
/// Resolves every existing path component through symbolic links/junctions and
/// preserves any not-yet-created tail below the resolved physical ancestor.
/// Storage-safety decisions must never rely on a user-facing alias path.
/// </summary>
public static class PhysicalPathResolver
{
    public static string Resolve(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("Path is empty.", nameof(path));
        }

        var full = Path.GetFullPath(path);
        var root = Path.GetPathRoot(full);
        if (string.IsNullOrWhiteSpace(root))
        {
            throw new IOException($"Unable to determine filesystem root for '{path}'.");
        }

        var current = root;
        var relative = Path.GetRelativePath(root, full);
        if (relative == ".") return PathSafety.Normalize(root);

        var segments = relative.Split(
            [Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar],
            StringSplitOptions.RemoveEmptyEntries);

        for (var index = 0; index < segments.Length; index++)
        {
            var candidate = Path.Combine(current, segments[index]);
            var isDirectory = Directory.Exists(candidate);
            var isFile = File.Exists(candidate);

            if (!isDirectory && !isFile)
            {
                // The destination commonly contains a not-yet-created event/year
                // tail. Once an ancestor is missing, no later component can be an
                // existing redirect, so append the remainder beneath the already
                // resolved physical ancestor.
                for (var tail = index; tail < segments.Length; tail++)
                {
                    current = Path.Combine(current, segments[tail]);
                }
                return PathSafety.Normalize(current);
            }

            FileSystemInfo info = isDirectory
                ? new DirectoryInfo(candidate)
                : new FileInfo(candidate);

            string? linkTarget;
            try
            {
                linkTarget = info.LinkTarget;
            }
            catch (Exception ex)
            {
                throw new IOException($"Unable to inspect path redirect '{candidate}'.", ex);
            }

            if (linkTarget is null)
            {
                current = Path.GetFullPath(candidate);
                continue;
            }

            FileSystemInfo? target;
            try
            {
                target = info.ResolveLinkTarget(returnFinalTarget: true);
            }
            catch (Exception ex)
            {
                throw new IOException($"Unable to resolve path redirect '{candidate}'.", ex);
            }

            if (target is null)
            {
                throw new IOException($"Path redirect '{candidate}' has no resolvable target.");
            }

            var resolvedTarget = Path.GetFullPath(target.FullName);
            if (!Directory.Exists(resolvedTarget) && !File.Exists(resolvedTarget))
            {
                throw new IOException($"Path redirect '{candidate}' targets an unavailable path.");
            }

            current = resolvedTarget;
        }

        return PathSafety.Normalize(current);
    }
}
