using System.IO;
using Microsoft.UI;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml.Navigation;
using WinRT.Interop;

namespace PandocGUI;

public partial class App : Application
{
    public static MainViewModel MainViewModel { get; } = new();

    public static Window MainWindow { get; private set; } = null!;

    private Window? window;

    public App()
    {
        InitializeComponent();
    }

    protected override void OnLaunched(LaunchActivatedEventArgs e)
    {
        window ??= new Window();
        MainWindow = window;

        if (window.Content is not Frame rootFrame)
        {
            rootFrame = new Frame();
            rootFrame.NavigationFailed += OnNavigationFailed;
            window.Content = rootFrame;
        }

        ConfigureWindow(window, rootFrame);
        _ = rootFrame.Navigate(typeof(MainPage), e.Arguments);
        window.Activate();
    }

    private static void ConfigureWindow(Window window, Frame rootFrame)
    {
        rootFrame.RequestedTheme = ElementTheme.Default;

        var hwnd = WindowNative.GetWindowHandle(window);
        var windowId = Win32Interop.GetWindowIdFromWindow(hwnd);
        var appWindow = AppWindow.GetFromWindowId(windowId);
        appWindow.Title = "PandocGUI";
        appWindow.SetIcon(Path.Combine(AppContext.BaseDirectory, "Assets", "LOGO.ico"));

        if (!AppWindowTitleBar.IsCustomizationSupported())
        {
            return;
        }

        var titleBar = appWindow.TitleBar;
        ApplyTitleBarTheme(titleBar, rootFrame.ActualTheme);
        titleBar.BackgroundColor = Colors.Transparent;
        titleBar.InactiveBackgroundColor = Colors.Transparent;
        titleBar.ButtonBackgroundColor = Colors.Transparent;
        titleBar.ButtonInactiveBackgroundColor = Colors.Transparent;
        rootFrame.ActualThemeChanged += (_, _) => ApplyTitleBarTheme(titleBar, rootFrame.ActualTheme);
    }

    private static void ApplyTitleBarTheme(AppWindowTitleBar titleBar, ElementTheme theme)
    {
        titleBar.PreferredTheme = theme == ElementTheme.Dark ? TitleBarTheme.Dark : TitleBarTheme.Light;
    }

    private void OnNavigationFailed(object sender, NavigationFailedEventArgs e)
    {
        throw new Exception("Failed to load Page " + e.SourcePageType.FullName);
    }
}
