namespace PandocGUI.Views;

public sealed partial class QueuePage : Page
{
    public QueuePage()
    {
        InitializeComponent();
    }

    public MainViewModel ViewModel => App.MainViewModel;
}
