using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace PhotoOrganizer.Core.Tests;

[TestClass]
public sealed class SafetyInvariantContractTests
{
    [TestMethod]
    public void ProductionCode_DestructiveFilesystemOperationsRemainExplicitlyAllowlisted()
    {
        var root = FindRepositoryRoot();
        var sourceRoot = Path.Combine(root, "src");
        var violations = new List<string>();

        foreach (var file in Directory.EnumerateFiles(sourceRoot, "*.cs", SearchOption.AllDirectories))
        {
            var relative = Relative(root, file);
            var lines = File.ReadAllLines(file);

            for (var index = 0; index < lines.Length; index++)
            {
                var line = lines[index];
                var lineNumber = index + 1;

                if (line.Contains("Directory.Delete(", StringComparison.Ordinal))
                {
                    violations.Add($"{relative}:{lineNumber}: production Directory.Delete is forbidden");
                }

                if (line.Contains("File.Delete(", StringComparison.Ordinal)
                    && !IsAllowedDelete(relative, line))
                {
                    violations.Add($"{relative}:{lineNumber}: unreviewed File.Delete: {line.Trim()}");
                }

                if (line.Contains("File.Move(", StringComparison.Ordinal)
                    && !IsAllowedMove(relative, line))
                {
                    violations.Add($"{relative}:{lineNumber}: unreviewed File.Move: {line.Trim()}");
                }

                if (line.Contains("File.Replace(", StringComparison.Ordinal))
                {
                    violations.Add($"{relative}:{lineNumber}: File.Replace is not part of the safety contract");
                }

                if (line.Contains("File.Copy(", StringComparison.Ordinal))
                {
                    violations.Add($"{relative}:{lineNumber}: direct File.Copy bypasses SafeCopyService");
                }
            }
        }

        Assert.AreEqual(
            0,
            violations.Count,
            "Production destructive-file API surface changed. Review the source-data safety contract before changing the allowlist:\n"
            + string.Join(Environment.NewLine, violations));
    }

    [TestMethod]
    public void SafeCopyService_RemainsNoClobberAndDeletesOnlyItsOwnedPartial()
    {
        var root = FindRepositoryRoot();
        var safeCopy = File.ReadAllText(Path.Combine(root, "src", "PhotoOrganizer.Core", "SafeCopyService.cs"));
        var durability = File.ReadAllText(Path.Combine(root, "src", "PhotoOrganizer.Core", "FileDurability.cs"));

        StringAssert.Contains(safeCopy, "FileMode.CreateNew");
        Assert.IsFalse(
            safeCopy.Contains("FileMode.Create,", StringComparison.Ordinal),
            "SafeCopyService must not create/overwrite a destination with FileMode.Create.");
        StringAssert.Contains(safeCopy, "File.Delete(temporaryPathForCleanup)");
        StringAssert.Contains(safeCopy, "temporaryPathForCleanup = null");
        StringAssert.Contains(durability, "File.Move(temporaryPath, finalPath, overwrite: false)");
        Assert.IsFalse(
            durability.Contains("MoveFileReplaceExisting", StringComparison.OrdinalIgnoreCase),
            "Windows finalization must not gain replace-existing semantics.");
    }

    [TestMethod]
    public void SafeToReuse_HasOneProductionDecisionPathAfterFinalConsistencyProof()
    {
        var root = FindRepositoryRoot();
        var productionFiles = Directory.EnumerateFiles(Path.Combine(root, "src"), "*.cs", SearchOption.AllDirectories).ToArray();
        var decisionOccurrences = new List<string>();

        foreach (var file in productionFiles)
        {
            var relative = Relative(root, file);
            if (relative == "src/PhotoOrganizer.Core/ImportModels.cs")
            {
                continue; // The model exposes IsSafeToReuse; it does not grant approval.
            }

            var lines = File.ReadAllLines(file);
            for (var index = 0; index < lines.Length; index++)
            {
                if (lines[index].Contains("ImportSafetyStatus.SafeToReuse", StringComparison.Ordinal))
                {
                    decisionOccurrences.Add($"{relative}:{index + 1}");
                }
            }
        }

        Assert.AreEqual(
            1,
            decisionOccurrences.Count,
            "Exactly one production code path may construct SafeToReuse. Found: " + string.Join(", ", decisionOccurrences));
        StringAssert.StartsWith(decisionOccurrences[0], "src/PhotoOrganizer.Core/ImportCoordinator.cs:");

        var coordinator = File.ReadAllText(Path.Combine(root, "src", "PhotoOrganizer.Core", "ImportCoordinator.cs"));
        var finalRescan = coordinator.IndexOf("var finalRescan = _scanner.Scan", StringComparison.Ordinal);
        var setEquality = coordinator.IndexOf("rescanned.SetEquals(finalFiles)", StringComparison.Ordinal);
        var safeDecision = coordinator.IndexOf("ImportSafetyStatus.SafeToReuse", StringComparison.Ordinal);

        Assert.IsTrue(finalRescan >= 0, "Final whole-card consistency scan disappeared.");
        Assert.IsTrue(setEquality > finalRescan, "Final supported-file set equality check must follow the final scan.");
        Assert.IsTrue(safeDecision > setEquality, "SafeToReuse must remain after the final consistency proof.");
    }

    [TestMethod]
    public void GreenReuseUi_HasOneLiteralAndIsGuardedByIsSafeToReuse()
    {
        var root = FindRepositoryRoot();
        const string greenHeadline = "保存先コピー検証済み — SDカード再利用可能";
        var occurrences = new List<string>();

        foreach (var file in Directory.EnumerateFiles(Path.Combine(root, "src"), "*.cs", SearchOption.AllDirectories))
        {
            var relative = Relative(root, file);
            var lines = File.ReadAllLines(file);
            for (var index = 0; index < lines.Length; index++)
            {
                if (lines[index].Contains(greenHeadline, StringComparison.Ordinal))
                {
                    occurrences.Add($"{relative}:{index + 1}");
                }
            }
        }

        Assert.AreEqual(1, occurrences.Count, "The green reuse-approved UI must have exactly one production assignment site.");
        StringAssert.StartsWith(occurrences[0], "src/PhotoOrganizer.App/MainWindowViewModel.ImportWorkflow.cs:");

        var workflow = File.ReadAllText(Path.Combine(root, "src", "PhotoOrganizer.App", "MainWindowViewModel.ImportWorkflow.cs"));
        var guard = workflow.IndexOf("if (result.IsSafeToReuse)", StringComparison.Ordinal);
        var green = workflow.IndexOf(greenHeadline, StringComparison.Ordinal);
        Assert.IsTrue(guard >= 0 && green > guard, "Green reuse approval must remain behind result.IsSafeToReuse.");
    }

    private static bool IsAllowedDelete(string relative, string line) =>
        (relative == "src/PhotoOrganizer.Core/SafeCopyService.cs"
         && line.Contains("File.Delete(temporaryPathForCleanup)", StringComparison.Ordinal))
        || (relative == "src/PhotoOrganizer.Core/SafeCopyService.cs"
            && line.Contains("File.Delete(temporaryPath)", StringComparison.Ordinal))
        || (relative == "src/PhotoOrganizer.App/JsonSettingsFile.cs"
            && line.Contains("File.Delete(temporary)", StringComparison.Ordinal))
        || (relative == "src/PhotoOrganizer.App/StartupRegistrationService.cs"
            && line.Contains("File.Delete(plistPath)", StringComparison.Ordinal));

    private static bool IsAllowedMove(string relative, string line) =>
        (relative == "src/PhotoOrganizer.Core/FileDurability.cs"
         && line.Contains("File.Move(temporaryPath, finalPath, overwrite: false)", StringComparison.Ordinal))
        || (relative == "src/PhotoOrganizer.App/JsonSettingsFile.cs"
            && line.Contains("File.Move(temporary, path, overwrite: true)", StringComparison.Ordinal))
        || (relative == "src/PhotoOrganizer.App/StartupRegistrationService.cs"
            && line.Contains("File.Move(temporary, plistPath, overwrite: true)", StringComparison.Ordinal));

    private static string FindRepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "PhotoOrganizer.slnx")))
            {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        // dotnet test normally runs below the checkout, but keep a working-directory
        // fallback for IDE/test-runner layouts that copy binaries elsewhere.
        directory = new DirectoryInfo(Environment.CurrentDirectory);
        while (directory is not null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "PhotoOrganizer.slnx")))
            {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        Assert.Fail("Could not locate repository root containing PhotoOrganizer.slnx.");
        throw new InvalidOperationException("Unreachable after Assert.Fail.");
    }

    private static string Relative(string root, string path) =>
        Path.GetRelativePath(root, path).Replace('\\', '/');
}
