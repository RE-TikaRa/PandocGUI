using System.Collections.ObjectModel;
using System.IO;
using Microsoft.UI.Xaml;
using PandocGUI.Models;
using PandocGUI.Services;
using PandocGUI.Utilities;
using Windows.Storage.Pickers;
using Windows.System;
using WinRT.Interop;

namespace PandocGUI.ViewModels;

public partial class MainViewModel : BaseViewModel
{
    private CancellationTokenSource? conversionCts;
    private CancellationTokenSource? downloadCts;
    private bool initialized;
    private string lastOutputFormat = "docx";

    public MainViewModel()
    {
        Title = "Pandoc GUI";
        Queue.CollectionChanged += (_, _) =>
        {
            StartConversionCommand.NotifyCanExecuteChanged();
            OnPropertyChanged(nameof(QueueCount));
        };
    }

    public ObservableCollection<ConversionItem> Queue { get; } = new();

    public ObservableCollection<string> InputFormats { get; } = new();

    public ObservableCollection<string> OutputFormats { get; } = new();

    private bool isBusy;
    private bool isPandocReady;
    private string pandocPath = string.Empty;
    private string pandocStatus = "未检测";
    private string selectedInputFormat = "Auto";
    private string selectedOutputFormat = "docx";
    private string outputDirectory = string.Empty;
    private string outputExtension = "docx";
    private string additionalArgs = string.Empty;
    private string templatePath = string.Empty;
    private string presetName = string.Empty;
    private string logText = string.Empty;
    private ConversionItem? selectedItem;
    private bool isDownloading;
    private double downloadProgress;
    private OutputPreset? selectedPreset;
    private string? selectedRecentFile;
    private readonly ObservableCollection<OutputPreset> presets = new();
    private readonly ObservableCollection<string> recentFiles = new();
    private List<OutputPreset> customPresets = new();

    public bool IsBusy
    {
        get => isBusy;
        set
        {
            if (SetProperty(ref isBusy, value))
            {
                OnIsBusyChanged(value);
            }
        }
    }

    public bool IsPandocReady
    {
        get => isPandocReady;
        set
        {
            if (SetProperty(ref isPandocReady, value))
            {
                OnIsPandocReadyChanged(value);
            }
        }
    }

    public string PandocPath
    {
        get => pandocPath;
        set
        {
            if (SetProperty(ref pandocPath, value))
            {
                OnPandocPathChanged(value);
            }
        }
    }

    public string PandocStatus
    {
        get => pandocStatus;
        set => SetProperty(ref pandocStatus, value);
    }

    public string SelectedInputFormat
    {
        get => selectedInputFormat;
        set
        {
            if (SetProperty(ref selectedInputFormat, value))
            {
                OnSelectedInputFormatChanged(value);
            }
        }
    }

    public string SelectedOutputFormat
    {
        get => selectedOutputFormat;
        set
        {
            if (SetProperty(ref selectedOutputFormat, value))
            {
                OnSelectedOutputFormatChanged(value);
            }
        }
    }

    public string OutputDirectory
    {
        get => outputDirectory;
        set
        {
            if (SetProperty(ref outputDirectory, value))
            {
                OnOutputDirectoryChanged(value);
            }
        }
    }

    public string OutputExtension
    {
        get => outputExtension;
        set
        {
            if (SetProperty(ref outputExtension, value))
            {
                OnOutputExtensionChanged(value);
            }
        }
    }

    public string AdditionalArgs
    {
        get => additionalArgs;
        set
        {
            if (SetProperty(ref additionalArgs, value))
            {
                OnAdditionalArgsChanged(value);
            }
        }
    }

    public string LogText
    {
        get => logText;
        set => SetProperty(ref logText, value);
    }

    public ConversionItem? SelectedItem
    {
        get => selectedItem;
        set => SetProperty(ref selectedItem, value);
    }

    public bool IsDownloading
    {
        get => isDownloading;
        set
        {
            if (SetProperty(ref isDownloading, value))
            {
                OnIsDownloadingChanged(value);
            }
        }
    }

    public double DownloadProgress
    {
        get => downloadProgress;
        set => SetProperty(ref downloadProgress, value);
    }

    public Visibility BusyVisibility => IsBusy ? Visibility.Visible : Visibility.Collapsed;

    public Visibility DownloadVisibility => IsDownloading ? Visibility.Visible : Visibility.Collapsed;

    public int QueueCount => Queue.Count;

    public async Task InitializeAsync(XamlRoot? root)
    {
        if (initialized)
        {
            return;
        }

        initialized = true;
        OutputDirectory = AppSettings.OutputDirectory ?? string.Empty;
        AdditionalArgs = AppSettings.AdditionalArgs ?? string.Empty;
        SelectedOutputFormat = AppSettings.SelectedOutputFormat ?? "docx";
        lastOutputFormat = SelectedOutputFormat;
        SelectedInputFormat = AppSettings.SelectedInputFormat ?? "Auto";
        OutputExtension = AppSettings.OutputExtension ?? SelectedOutputFormat;
        PandocPath = AppSettings.PandocPath ?? string.Empty;
        if (!AppSettings.HasLaunchedBefore)
        {
            AppSettings.HasLaunchedBefore = true;
        }

        await RefreshPandocAsync();
    }

    [RelayCommand]
    private async Task RefreshPandocAsync()
    {
        AppendLog("开始检测 pandoc...");
        PandocStatus = "正在检测...";
        var detection = PandocService.DetectPandoc(PandocPath);
        foreach (var step in detection.Steps)
        {
            AppendLog(step);
        }

        var path = detection.PandocPath;
        if (string.IsNullOrWhiteSpace(path))
        {
            IsPandocReady = false;
            PandocStatus = "未检测到 pandoc";
            AppendLog("未检测到 pandoc，请检查安装路径或点击“选择”。");
            InputFormats.Clear();
            OutputFormats.Clear();
            return;
        }

        PandocPath = path;
        AppendLog($"最终选择：{path}（{detection.Source}）");
        var info = await PandocService.GetInfoAsync(path, CancellationToken.None);
        if (info is null)
        {
            IsPandocReady = false;
            PandocStatus = "已找到但不可用";
            AppendLog("pandoc 已找到但不可用，请确认能在命令行运行。");
            return;
        }

        IsPandocReady = true;
        PandocStatus = $"已就绪 {info.Version}";
        AppendLog($"pandoc 版本: {info.Version}");
        await LoadFormatsAsync(path);
    }

    [RelayCommand]
    private async Task SelectPandocAsync()
    {
        var picker = new FileOpenPicker();
        picker.FileTypeFilter.Add(".exe");
        InitializePicker(picker);

        var file = await picker.PickSingleFileAsync();
        if (file is null)
        {
            return;
        }

        PandocPath = file.Path;
        await RefreshPandocAsync();
    }

    [RelayCommand]
    private async Task DownloadPandocAsync()
    {
        if (IsDownloading)
        {
            return;
        }

        IsDownloading = true;
        DownloadProgress = 0;
        downloadCts?.Cancel();
        downloadCts = new CancellationTokenSource();

        try
        {
            var progress = new Progress<double>(value => DownloadProgress = value);
            var result = await PandocService.DownloadLatestAsync(progress, downloadCts.Token);
            if (result.Succeeded && !string.IsNullOrWhiteSpace(result.PandocPath))
            {
                PandocPath = result.PandocPath;
                AppendLog($"已下载 Pandoc: {result.PandocPath}");
                await RefreshPandocAsync();
            }
            else
            {
                AppendLog(result.ErrorMessage ?? "下载失败");
            }
        }
        catch (OperationCanceledException)
        {
            AppendLog("下载已取消");
        }
        catch (Exception ex)
        {
            AppendLog(ex.Message);
        }
        finally
        {
            IsDownloading = false;
        }
    }

    [RelayCommand]
    private async Task OpenPandocWebsiteAsync()
    {
        await Launcher.LaunchUriAsync(new Uri("https://pandoc.org/installing.html"));
    }

    [RelayCommand]
    private async Task AddFilesAsync()
    {
        var picker = new FileOpenPicker();
        picker.FileTypeFilter.Add("*");
        InitializePicker(picker);

        var files = await picker.PickMultipleFilesAsync();
        if (files is null || files.Count == 0)
        {
            return;
        }

        foreach (var file in files)
        {
            AddFile(file.Path);
        }
    }

    [RelayCommand]
    private async Task AddFolderAsync()
    {
        var picker = new FolderPicker();
        picker.FileTypeFilter.Add("*");
        InitializePicker(picker);

        var folder = await picker.PickSingleFolderAsync();
        if (folder is null)
        {
            return;
        }

        foreach (var path in Directory.GetFiles(folder.Path))
        {
            AddFile(path);
        }
    }

    [RelayCommand]
    private async Task BrowseOutputFolderAsync()
    {
        var picker = new FolderPicker();
        picker.FileTypeFilter.Add("*");
        InitializePicker(picker);

        var folder = await picker.PickSingleFolderAsync();
        if (folder is not null)
        {
            OutputDirectory = folder.Path;
        }
    }

    [RelayCommand]
    private void RemoveSelected()
    {
        if (SelectedItem is null)
        {
            return;
        }

        Queue.Remove(SelectedItem);
    }

    [RelayCommand]
    private void ClearQueue()
    {
        Queue.Clear();
    }

    [RelayCommand(CanExecute = nameof(CanStartConversion))]
    private async Task StartConversionAsync()
    {
        if (!IsPandocReady || string.IsNullOrWhiteSpace(PandocPath))
        {
            AppendLog("pandoc 未就绪，无法开始转换");
            return;
        }

        IsBusy = true;
        conversionCts?.Cancel();
        conversionCts = new CancellationTokenSource();

        try
        {
            foreach (var item in Queue)
            {
                if (conversionCts.IsCancellationRequested)
                {
                    item.Status = ConversionStatus.Skipped;
                    item.Message = "已取消";
                    continue;
                }

                item.Status = ConversionStatus.Running;
                item.Message = string.Empty;

                if (string.IsNullOrWhiteSpace(item.OutputPath))
                {
                    item.OutputPath = BuildOutputPath(item.InputPath);
                }

                var outputDirectory = Path.GetDirectoryName(item.OutputPath);
                if (!string.IsNullOrWhiteSpace(outputDirectory))
                {
                    Directory.CreateDirectory(outputDirectory);
                }

                var args = BuildArguments(item);
                var result = await PandocService.RunAsync(PandocPath, args, conversionCts.Token);

                if (result.Succeeded)
                {
                    item.Status = ConversionStatus.Succeeded;
                    item.Message = "完成";
                    AppendLog($"{item.FileName} -> {item.OutputPath}");
                }
                else
                {
                    item.Status = ConversionStatus.Failed;
                    item.Message = "失败";
                    AppendLog($"失败: {item.FileName}");
                }

                AppendLog(result.StandardError);
            }
        }
        catch (OperationCanceledException)
        {
            AppendLog("转换已取消");
        }
        catch (Exception ex)
        {
            AppendLog(ex.Message);
        }
        finally
        {
            IsBusy = false;
            conversionCts?.Dispose();
            conversionCts = null;
        }
    }

    private bool CanStartConversion()
        => !IsBusy && IsPandocReady && Queue.Count > 0;

    [RelayCommand(CanExecute = nameof(CanCancelConversion))]
    private void CancelConversion()
    {
        conversionCts?.Cancel();
    }

    private bool CanCancelConversion()
        => IsBusy;

    private void AddFile(string path)
    {
        if (Queue.Any(item => string.Equals(item.InputPath, path, StringComparison.OrdinalIgnoreCase)))
        {
            return;
        }

        if (string.IsNullOrWhiteSpace(OutputDirectory))
        {
            OutputDirectory = Path.GetDirectoryName(path) ?? string.Empty;
        }

        var item = new ConversionItem(path)
        {
            OutputPath = BuildOutputPath(path)
        };

        Queue.Add(item);
    }

    private string BuildOutputPath(string inputPath)
    {
        var directory = OutputDirectory;
        if (string.IsNullOrWhiteSpace(directory))
        {
            directory = Path.GetDirectoryName(inputPath) ?? string.Empty;
        }

        var extension = string.IsNullOrWhiteSpace(OutputExtension)
            ? SelectedOutputFormat
            : OutputExtension.TrimStart('.');

        var fileName = Path.GetFileNameWithoutExtension(inputPath);
        return Path.Combine(directory, $"{fileName}.{extension}");
    }

    private List<string> BuildArguments(ConversionItem item)
    {
        var args = new List<string>();

        if (!string.IsNullOrWhiteSpace(SelectedInputFormat) &&
            !string.Equals(SelectedInputFormat, "Auto", StringComparison.OrdinalIgnoreCase))
        {
            args.Add("-f");
            args.Add(SelectedInputFormat);
        }

        if (!string.IsNullOrWhiteSpace(SelectedOutputFormat))
        {
            args.Add("-t");
            args.Add(SelectedOutputFormat);
        }

        args.AddRange(CommandLineTokenizer.Split(AdditionalArgs));
        args.Add("-o");
        args.Add(item.OutputPath);
        args.Add(item.InputPath);

        return args;
    }

    private async Task LoadFormatsAsync(string pandocPath)
    {
        var inputFormats = await PandocService.ListInputFormatsAsync(pandocPath, CancellationToken.None);
        var outputFormats = await PandocService.ListOutputFormatsAsync(pandocPath, CancellationToken.None);

        InputFormats.Clear();
        InputFormats.Add("Auto");
        foreach (var format in inputFormats)
        {
            InputFormats.Add(format);
        }

        OutputFormats.Clear();
        foreach (var format in outputFormats)
        {
            OutputFormats.Add(format);
        }

        if (!InputFormats.Any(f => f.Equals(SelectedInputFormat, StringComparison.OrdinalIgnoreCase)))
        {
            SelectedInputFormat = "Auto";
        }

        if (!OutputFormats.Any(f => f.Equals(SelectedOutputFormat, StringComparison.OrdinalIgnoreCase)) &&
            OutputFormats.Count > 0)
        {
            SelectedOutputFormat = OutputFormats[0];
        }
    }

    private void AppendLog(string? message)
    {
        if (string.IsNullOrWhiteSpace(message))
        {
            return;
        }

        var line = message.TrimEnd();
        if (string.IsNullOrWhiteSpace(line))
        {
            return;
        }

        LogText += $"[{DateTime.Now:HH:mm:ss}] {line}{Environment.NewLine}";
    }

    private static void InitializePicker(object picker)
    {
        var hwnd = WindowNative.GetWindowHandle(App.MainWindow);
        InitializeWithWindow.Initialize(picker, hwnd);
    }

    private void OnIsBusyChanged(bool value)
    {
        OnPropertyChanged(nameof(BusyVisibility));
        StartConversionCommand.NotifyCanExecuteChanged();
        CancelConversionCommand.NotifyCanExecuteChanged();
    }

    private void OnIsPandocReadyChanged(bool value)
    {
        StartConversionCommand.NotifyCanExecuteChanged();
    }

    private void OnIsDownloadingChanged(bool value)
    {
        OnPropertyChanged(nameof(DownloadVisibility));
    }

    private void OnPandocPathChanged(string value)
    {
        AppSettings.PandocPath = value;
    }

    private void OnOutputDirectoryChanged(string value)
    {
        AppSettings.OutputDirectory = value;
        UpdateOutputPathsForPending();
    }

    private void OnAdditionalArgsChanged(string value)
    {
        AppSettings.AdditionalArgs = value;
    }

    private void OnSelectedOutputFormatChanged(string value)
    {
        AppSettings.SelectedOutputFormat = value;
        if (string.IsNullOrWhiteSpace(OutputExtension) ||
            OutputExtension.Equals(lastOutputFormat, StringComparison.OrdinalIgnoreCase))
        {
            OutputExtension = value;
        }

        lastOutputFormat = value;
    }

    private void OnSelectedInputFormatChanged(string value)
    {
        AppSettings.SelectedInputFormat = value;
    }

    private void OnOutputExtensionChanged(string value)
    {
        AppSettings.OutputExtension = value;
        UpdateOutputPathsForPending();
    }

    private void UpdateOutputPathsForPending()
    {
        foreach (var item in Queue)
        {
            if (item.Status == ConversionStatus.Pending)
            {
                item.OutputPath = BuildOutputPath(item.InputPath);
            }
        }
    }
}
