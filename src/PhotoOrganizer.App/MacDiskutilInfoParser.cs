using System.Xml.Linq;

namespace PhotoOrganizer.App;

internal sealed record MacStorageIdentity(
    string VolumeFingerprint,
    string? PhysicalDeviceFingerprint);

internal static class MacDiskutilInfoParser
{
    public static MacStorageIdentity? Parse(string? output)
    {
        if (string.IsNullOrWhiteSpace(output)) return null;

        try
        {
            var document = XDocument.Parse(output, LoadOptions.None);
            var dictionary = document.Root?.Name.LocalName == "dict"
                ? document.Root
                : document.Root?.Elements().FirstOrDefault(element => element.Name.LocalName == "dict");
            if (dictionary is null) return null;

            // Prefer persistent filesystem/partition identities. DeviceIdentifier is
            // intentionally the last fallback because BSD disk numbers are ephemeral.
            var persistentVolumeId = FirstNonEmpty(
                GetPlistString(dictionary, "VolumeUUID"),
                GetPlistString(dictionary, "APFSVolumeUUID"),
                GetPlistString(dictionary, "DiskUUID"),
                GetPlistString(dictionary, "MediaUUID"));
            var deviceIdentifier = Clean(GetPlistString(dictionary, "DeviceIdentifier"));
            var volumeIdentity = FirstNonEmpty(persistentVolumeId, deviceIdentifier);
            if (volumeIdentity is null) return null;

            var parentWholeDisk = Clean(GetPlistString(dictionary, "ParentWholeDisk"));
            var wholeDisk = GetPlistBoolean(dictionary, "WholeDisk") == true
                ? deviceIdentifier
                : null;
            var physicalIdentity = FirstNonEmpty(parentWholeDisk, wholeDisk);

            return new MacStorageIdentity(
                "mac-volume:" + volumeIdentity,
                physicalIdentity is null
                    ? null
                    : "mac-whole-disk:" + physicalIdentity);
        }
        catch
        {
            // Malformed or unexpected diskutil output is a failed identity proof.
            return null;
        }
    }

    private static string? FirstNonEmpty(params string?[] values)
    {
        foreach (var value in values)
        {
            var cleaned = Clean(value);
            if (cleaned is not null) return cleaned;
        }

        return null;
    }

    private static string? Clean(string? value) =>
        string.IsNullOrWhiteSpace(value) ? null : value.Trim();

    private static string? GetPlistString(XElement dictionary, string key)
    {
        var elements = dictionary.Elements().ToArray();
        for (var index = 0; index + 1 < elements.Length; index++)
        {
            if (elements[index].Name.LocalName != "key" || elements[index].Value != key) continue;
            return elements[index + 1].Name.LocalName == "string"
                ? elements[index + 1].Value
                : null;
        }

        return null;
    }

    private static bool? GetPlistBoolean(XElement dictionary, string key)
    {
        var elements = dictionary.Elements().ToArray();
        for (var index = 0; index + 1 < elements.Length; index++)
        {
            if (elements[index].Name.LocalName != "key" || elements[index].Value != key) continue;
            return elements[index + 1].Name.LocalName switch
            {
                "true" => true,
                "false" => false,
                _ => null
            };
        }

        return null;
    }
}
