namespace PandocGUI.Views;

public sealed partial class PresetsPage : Page
{
    public PresetsPage()
    {
        InitializeComponent();
    }

    public MainViewModel ViewModel => App.MainViewModel;
}
