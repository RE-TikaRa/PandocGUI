using System.Linq;

namespace PandocGUI.Views;

public sealed partial class RecentPage : Page
{
    public RecentPage()
    {
        InitializeComponent();
    }

    public MainViewModel ViewModel => App.MainViewModel;

    private void OnDragOver(object sender, DragEventArgs e)
    {
        e.AcceptedOperation = Windows.ApplicationModel.DataTransfer.DataPackageOperation.Copy;
        e.Handled = true;
    }

    private async void OnDrop(object sender, DragEventArgs e)
    {
        if (!e.DataView.Contains(Windows.ApplicationModel.DataTransfer.StandardDataFormats.StorageItems))
        {
            return;
        }

        var items = await e.DataView.GetStorageItemsAsync();
        var paths = items.Select(item => item.Path);
        ViewModel.AddFilesFromPaths(paths);
    }
}
