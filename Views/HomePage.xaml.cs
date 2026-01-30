using Microsoft.UI.Xaml.Media.Imaging;

namespace PandocGUI.Views;

public sealed partial class HomePage : Page
{
    public HomePage()
    {
        InitializeComponent();
        SizeChanged += OnPageSizeChanged;
        Loaded += (_, _) => UpdateLogoForTheme();
        ActualThemeChanged += (_, _) => UpdateLogoForTheme();
    }

    public MainViewModel ViewModel => App.MainViewModel;

    private void OnPageSizeChanged(object sender, SizeChangedEventArgs e)
    {
        var size = Math.Min(e.NewSize.Width, e.NewSize.Height) * 0.6;
        LogoImage.Width = size;
        LogoImage.Height = size;
    }

    private void UpdateLogoForTheme()
    {
        var isDark = ActualTheme == ElementTheme.Dark;
        var uri = new Uri(isDark ? "ms-appx:///Assets/LOGO_dark.svg" : "ms-appx:///Assets/LOGO.svg");
        LogoImage.Source = new SvgImageSource(uri);
    }
}
