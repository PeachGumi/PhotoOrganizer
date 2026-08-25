using System.Xml.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace PhotoOrganizer.App.Tests;

[TestClass]
public sealed class StartupRegistrationFormatterTests
{
    [TestMethod]
    public void WindowsCommand_AlwaysQuotesExecutablePath()
    {
        var command = StartupRegistrationFormatter.BuildWindowsCommand(
            @"C:\Program Files\Photo Organizer\PhotoOrganizer.exe",
            startInBackground: false);

        Assert.AreEqual(
            "\"C:\\Program Files\\Photo Organizer\\PhotoOrganizer.exe\"",
            command);
    }

    [TestMethod]
    public void WindowsCommand_BackgroundFlagIsExplicit()
    {
        var foreground = StartupRegistrationFormatter.BuildWindowsCommand(
            @"C:\PhotoOrganizer.exe",
            startInBackground: false);
        var background = StartupRegistrationFormatter.BuildWindowsCommand(
            @"C:\PhotoOrganizer.exe",
            startInBackground: true);

        Assert.IsFalse(foreground.Contains("--background", StringComparison.Ordinal));
        Assert.AreEqual("\"C:\\PhotoOrganizer.exe\" --background", background);
    }

    [TestMethod]
    public void WindowsCommand_RejectsBlankExecutable()
    {
        Assert.ThrowsExactly<ArgumentException>(() =>
            StartupRegistrationFormatter.BuildWindowsCommand("   ", false));
    }

    [TestMethod]
    public void MacPlist_IsValidXmlWithExpectedLabelAndExecutable()
    {
        var plist = StartupRegistrationFormatter.BuildMacLaunchAgentPlist(
            "com.peachgumi.photoorganizer",
            "/Applications/Photo Organizer.app/Contents/MacOS/PhotoOrganizer",
            startInBackground: false);

        var document = XDocument.Parse(plist, LoadOptions.None);
        var strings = document.Descendants().Where(e => e.Name.LocalName == "string").Select(e => e.Value).ToArray();

        CollectionAssert.Contains(strings, "com.peachgumi.photoorganizer");
        CollectionAssert.Contains(strings, "/Applications/Photo Organizer.app/Contents/MacOS/PhotoOrganizer");
    }

    [TestMethod]
    public void MacPlist_XmlEscapesSpecialCharactersWithoutChangingParsedValues()
    {
        const string label = "com.example.photo&organizer";
        const string executable = "/Applications/A&B <Photo>.app/Contents/MacOS/PhotoOrganizer";

        var plist = StartupRegistrationFormatter.BuildMacLaunchAgentPlist(label, executable, false);
        var document = XDocument.Parse(plist, LoadOptions.None);
        var strings = document.Descendants().Where(e => e.Name.LocalName == "string").Select(e => e.Value).ToArray();

        CollectionAssert.Contains(strings, label);
        CollectionAssert.Contains(strings, executable);
        StringAssert.Contains(plist, "&amp;");
        StringAssert.Contains(plist, "&lt;");
    }

    [TestMethod]
    public void MacPlist_BackgroundArgumentIsIncludedOnlyWhenRequested()
    {
        var foreground = StartupRegistrationFormatter.BuildMacLaunchAgentPlist(
            "com.example.foreground",
            "/tmp/PhotoOrganizer",
            false);
        var background = StartupRegistrationFormatter.BuildMacLaunchAgentPlist(
            "com.example.background",
            "/tmp/PhotoOrganizer",
            true);

        Assert.IsFalse(foreground.Contains("--background", StringComparison.Ordinal));
        Assert.AreEqual(1, CountStringValue(background, "--background"));
    }

    [TestMethod]
    public void MacPlist_RunAtLoadIsTrue()
    {
        var plist = StartupRegistrationFormatter.BuildMacLaunchAgentPlist(
            "com.example.app",
            "/tmp/PhotoOrganizer",
            false);
        var document = XDocument.Parse(plist, LoadOptions.None);
        var elements = document.Descendants().ToArray();
        var runAtLoadKey = Array.FindIndex(
            elements,
            e => e.Name.LocalName == "key" && e.Value == "RunAtLoad");

        Assert.IsTrue(runAtLoadKey >= 0);
        Assert.IsTrue(runAtLoadKey + 1 < elements.Length);
        Assert.AreEqual("true", elements[runAtLoadKey + 1].Name.LocalName);
    }

    [TestMethod]
    public void MacPlist_RejectsBlankLabelOrExecutable()
    {
        Assert.ThrowsExactly<ArgumentException>(() =>
            StartupRegistrationFormatter.BuildMacLaunchAgentPlist(" ", "/tmp/app", false));
        Assert.ThrowsExactly<ArgumentException>(() =>
            StartupRegistrationFormatter.BuildMacLaunchAgentPlist("com.example.app", " ", false));
    }

    private static int CountStringValue(string plist, string value)
    {
        var document = XDocument.Parse(plist, LoadOptions.None);
        return document.Descendants()
            .Count(element => element.Name.LocalName == "string" && element.Value == value);
    }
}
