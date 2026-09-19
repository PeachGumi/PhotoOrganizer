using System.Reflection;
using System.Runtime.CompilerServices;
using Avalonia.Controls.ApplicationLifetimes;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PhotoOrganizer.App;

namespace PhotoOrganizer.App.Tests;

[TestClass]
public sealed class ApplicationShutdownTests
{
    [TestMethod]
    public void IdlePlatformShutdown_AllowsHiddenMainWindowToClose()
    {
        var app = (App)RuntimeHelpers.GetUninitializedObject(typeof(App));
        var mainWindow = (MainWindow)RuntimeHelpers.GetUninitializedObject(typeof(MainWindow));

        SetPrivateField(app, "_mainWindow", mainWindow);
        var shutdownRequested = typeof(App).GetMethod(
            "OnShutdownRequested",
            BindingFlags.Instance | BindingFlags.NonPublic);
        Assert.IsNotNull(shutdownRequested);

        var eventArgs = new ShutdownRequestedEventArgs();
        shutdownRequested.Invoke(app, [null, eventArgs]);

        Assert.IsFalse(eventArgs.Cancel);
        Assert.IsTrue(GetPrivateField<bool>(mainWindow, "_allowExplicitClose"));
    }

    private static void SetPrivateField<T>(T instance, string name, object? value)
    {
        var field = typeof(T).GetField(name, BindingFlags.Instance | BindingFlags.NonPublic);
        Assert.IsNotNull(field);
        field.SetValue(instance, value);
    }

    private static TValue GetPrivateField<TValue>(object instance, string name)
    {
        var field = instance.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic);
        Assert.IsNotNull(field);
        return (TValue)field.GetValue(instance)!;
    }
}
