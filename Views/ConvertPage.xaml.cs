namespace PandocGUI.Views;

public sealed partial class ConvertPage : Page
{
    public ConvertPage()
    {
        InitializeComponent();
    }

    public MainViewModel ViewModel => App.MainViewModel;
}
