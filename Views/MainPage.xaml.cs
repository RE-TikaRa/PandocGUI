namespace PandocGUI.Views;

public partial class MainPage : Page
{
    public MainPage()
    {
        InitializeComponent();
        Loaded += OnLoaded;
    }

    public MainViewModel ViewModel => App.MainViewModel;

    private async void OnLoaded(object sender, RoutedEventArgs e)
    {
        App.MainWindow.ExtendsContentIntoTitleBar = true;
        App.MainWindow.SetTitleBar(AppTitleBar);
        await ViewModel.InitializeAsync(XamlRoot);
        if (ContentFrame.CurrentSourcePageType is null)
        {
            ContentFrame.Navigate(typeof(HomePage));
        }
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
            "convert-settings" => typeof(ConvertSettingsPage),
            "presets" => typeof(PresetsPage),
            "recent" => typeof(RecentPage),
            "queue" => typeof(QueuePage),
            "settings" => typeof(SettingsPage),
            _ => typeof(HomePage)
        };

        if (ContentFrame.CurrentSourcePageType != target)
        {
            ContentFrame.Navigate(target);
        }
    }
}
