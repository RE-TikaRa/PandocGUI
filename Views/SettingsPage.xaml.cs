namespace PandocGUI.Views;

public sealed partial class SettingsPage : Page
{
    public SettingsPage()
    {
        InitializeComponent();
    }

    public MainViewModel ViewModel => App.MainViewModel;
}
