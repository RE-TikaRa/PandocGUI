namespace PandocGUI.Views;

public sealed partial class OutputFormatPage : Page
{
    public OutputFormatPage()
    {
        InitializeComponent();
    }

    public MainViewModel ViewModel => App.MainViewModel;
}
