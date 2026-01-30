namespace PandocGUI.Views;

public sealed partial class ConvertSettingsPage : Page
{
    public ConvertSettingsPage()
    {
        InitializeComponent();
    }

    public MainViewModel ViewModel => App.MainViewModel;
}
