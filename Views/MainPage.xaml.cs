using System;
using System.Linq;

namespace PandocGUI.Views;

public partial class MainPage : Page
{
    public MainPage()
    {
        InitializeComponent();
        Loaded += OnLoaded;
    }

    public MainViewModel ViewModel => App.MainViewModel;

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        App.MainWindow.ExtendsContentIntoTitleBar = true;
        App.MainWindow.SetTitleBar(AppTitleBar);
        if (ContentFrame.CurrentSourcePageType is null)
        {
            NavigateToTag("home");
        }

        _ = ViewModel.InitializeAsync(XamlRoot);
    }

    private void OnNavigationSelectionChanged(NavigationView sender, NavigationViewSelectionChangedEventArgs args)
    {
        if (args.SelectedItem is not NavigationViewItem item)
        {
            return;
        }

        var tag = item.Tag?.ToString();
        var target = tag switch
        {
            "home" => typeof(HomePage),
            "convert" => typeof(ConvertPage),
            "settings" => typeof(SettingsPage),
            _ => typeof(HomePage)
        };

        if (ContentFrame.CurrentSourcePageType != target)
        {
            ContentFrame.Navigate(target);
        }
    }

    private void NavigateToTag(string tag)
    {
        var item = RootNavigationView.MenuItems
            .OfType<NavigationViewItem>()
            .FirstOrDefault(navItem => string.Equals(navItem.Tag?.ToString(), tag, StringComparison.Ordinal));
        if (item is not null)
        {
            RootNavigationView.SelectedItem = item;
        }

        var target = tag switch
        {
            "home" => typeof(HomePage),
            "convert" => typeof(ConvertPage),
            "settings" => typeof(SettingsPage),
            _ => typeof(HomePage)
        };

        if (ContentFrame.CurrentSourcePageType != target)
        {
            ContentFrame.Navigate(target);
        }
    }
}
