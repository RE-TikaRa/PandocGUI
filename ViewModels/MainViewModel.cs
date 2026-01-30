using System.Collections.ObjectModel;
using System.IO;
using System.Text.Json;
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
            RetryFailedCommand.NotifyCanExecuteChanged();
            ClearCompletedCommand.NotifyCanExecuteChanged();
        };
    }

    public ObservableCollection<ConversionItem> Queue { get; } = new();

    public ObservableCollection<string> InputFormats { get; } = new();

    public ObservableCollection<string> OutputFormats { get; } = new();

    private bool isBusy;
    private bool isPandocReady;
    private bool isQueuePaused;
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
    private TaskCompletionSource<bool>? pauseTcs;
    private readonly object pauseLock = new();
    private int maxParallelism = 1;
    private string? selectedRecentOutputDirectory;
    private string? selectedRecentTemplate;
    private OutputPreset? selectedPreset;
    private string? selectedRecentFile;
    private readonly ObservableCollection<OutputPreset> presets = new();
    private readonly ObservableCollection<string> recentFiles = new();
    private readonly ObservableCollection<string> recentOutputDirectories = new();
    private readonly ObservableCollection<string> recentTemplates = new();
    private readonly ObservableCollection<int> parallelismOptions = new() { 1, 2, 3, 4, 5, 6, 7, 8 };
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

    public bool IsQueuePaused
    {
        get => isQueuePaused;
        set
        {
            if (SetProperty(ref isQueuePaused, value))
            {
                PauseQueueCommand.NotifyCanExecuteChanged();
                ResumeQueueCommand.NotifyCanExecuteChanged();
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

    public ObservableCollection<int> ParallelismOptions => parallelismOptions;

    public int MaxParallelism
    {
        get => maxParallelism;
        set
        {
            var normalized = Math.Clamp(value, 1, 8);
            if (SetProperty(ref maxParallelism, normalized))
            {
                AppSettings.MaxParallelism = normalized;
            }
        }
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

    public string TemplatePath
    {
        get => templatePath;
        set
        {
            if (SetProperty(ref templatePath, value))
            {
                OnTemplatePathChanged(value);
            }
        }
    }

    public ObservableCollection<OutputPreset> Presets => presets;

    public OutputPreset? SelectedPreset
    {
        get => selectedPreset;
        set
        {
            if (SetProperty(ref selectedPreset, value))
            {
                OnSelectedPresetChanged();
            }
        }
    }

    public string PresetName
    {
        get => presetName;
        set
        {
            if (SetProperty(ref presetName, value))
            {
                SavePresetCommand.NotifyCanExecuteChanged();
            }
        }
    }

    public ObservableCollection<string> RecentFiles => recentFiles;

    public string? SelectedRecentFile
    {
        get => selectedRecentFile;
        set
        {
            if (SetProperty(ref selectedRecentFile, value))
            {
                AddRecentSelectedCommand.NotifyCanExecuteChanged();
            }
        }
    }

    public ObservableCollection<string> RecentOutputDirectories => recentOutputDirectories;

    public string? SelectedRecentOutputDirectory
    {
        get => selectedRecentOutputDirectory;
        set
        {
            if (SetProperty(ref selectedRecentOutputDirectory, value))
            {
                ApplyRecentOutputDirectoryCommand.NotifyCanExecuteChanged();
            }
        }
    }

    public ObservableCollection<string> RecentTemplates => recentTemplates;

    public string? SelectedRecentTemplate
    {
        get => selectedRecentTemplate;
        set
        {
            if (SetProperty(ref selectedRecentTemplate, value))
            {
                ApplyRecentTemplateCommand.NotifyCanExecuteChanged();
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
        set
        {
            if (SetProperty(ref selectedItem, value))
            {
                OpenOutputFolderCommand.NotifyCanExecuteChanged();
            }
        }
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
        TemplatePath = AppSettings.TemplatePath ?? string.Empty;
        MaxParallelism = AppSettings.MaxParallelism;
        LoadPresets();
        LoadRecentFiles();
        LoadRecentOutputDirectories();
        LoadRecentTemplates();
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
    private async Task BrowseTemplateAsync()
    {
        var picker = new FileOpenPicker();
        picker.FileTypeFilter.Add("*");
        InitializePicker(picker);

        var file = await picker.PickSingleFileAsync();
        if (file is not null)
        {
            TemplatePath = file.Path;
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

    [RelayCommand(CanExecute = nameof(CanRetryFailed))]
    private void RetryFailed()
    {
        foreach (var item in Queue.Where(item => item.Status == ConversionStatus.Failed))
        {
            item.Status = ConversionStatus.Pending;
            item.Message = string.Empty;
        }
    }

    private bool CanRetryFailed()
        => !IsBusy && Queue.Any(item => item.Status == ConversionStatus.Failed);

    [RelayCommand(CanExecute = nameof(CanClearCompleted))]
    private void ClearCompleted()
    {
        var completed = Queue.Where(item => item.Status != ConversionStatus.Pending && item.Status != ConversionStatus.Running)
            .ToList();
        foreach (var item in completed)
        {
            Queue.Remove(item);
        }
    }

    private bool CanClearCompleted()
        => Queue.Any(item => item.Status != ConversionStatus.Pending && item.Status != ConversionStatus.Running);

    [RelayCommand(CanExecute = nameof(CanOpenOutputFolder))]
    private async Task OpenOutputFolderAsync()
    {
        if (SelectedItem is null || string.IsNullOrWhiteSpace(SelectedItem.OutputPath))
        {
            return;
        }

        var directory = Path.GetDirectoryName(SelectedItem.OutputPath);
        if (string.IsNullOrWhiteSpace(directory))
        {
            return;
        }

        await Launcher.LaunchFolderPathAsync(directory);
    }

    private bool CanOpenOutputFolder()
        => SelectedItem is not null && !string.IsNullOrWhiteSpace(SelectedItem.OutputPath);

    [RelayCommand(CanExecute = nameof(CanApplyPreset))]
    private void ApplyPreset()
    {
        if (SelectedPreset is null)
        {
            return;
        }

        if (!string.IsNullOrWhiteSpace(SelectedPreset.OutputFormat))
        {
            SelectedOutputFormat = SelectedPreset.OutputFormat;
        }

        if (!string.IsNullOrWhiteSpace(SelectedPreset.OutputExtension))
        {
            OutputExtension = SelectedPreset.OutputExtension;
        }

        AdditionalArgs = SelectedPreset.AdditionalArgs ?? string.Empty;

        if (!string.IsNullOrWhiteSpace(SelectedPreset.TemplatePath))
        {
            TemplatePath = SelectedPreset.TemplatePath;
        }
    }

    private bool CanApplyPreset()
        => SelectedPreset is not null;

    [RelayCommand(CanExecute = nameof(CanSavePreset))]
    private void SavePreset()
    {
        var name = PresetName.Trim();
        if (string.IsNullOrWhiteSpace(name))
        {
            return;
        }

        var preset = new OutputPreset
        {
            Name = name,
            OutputFormat = SelectedOutputFormat,
            OutputExtension = OutputExtension,
            AdditionalArgs = AdditionalArgs,
            TemplatePath = TemplatePath,
            IsBuiltIn = false
        };

        var existing = customPresets.FirstOrDefault(p => p.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
        if (existing is null)
        {
            customPresets.Add(preset);
        }
        else
        {
            existing.OutputFormat = preset.OutputFormat;
            existing.OutputExtension = preset.OutputExtension;
            existing.AdditionalArgs = preset.AdditionalArgs;
            existing.TemplatePath = preset.TemplatePath;
        }

        AppSettings.SetPresets(customPresets);
        LoadPresets();
        SelectedPreset = Presets.FirstOrDefault(p => p.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
    }

    private bool CanSavePreset()
        => !string.IsNullOrWhiteSpace(PresetName);

    [RelayCommand(CanExecute = nameof(CanRemovePreset))]
    private void RemovePreset()
    {
        if (SelectedPreset is null || SelectedPreset.IsBuiltIn)
        {
            return;
        }

        customPresets = customPresets
            .Where(p => !p.Name.Equals(SelectedPreset.Name, StringComparison.OrdinalIgnoreCase))
            .ToList();
        AppSettings.SetPresets(customPresets);
        LoadPresets();
    }

    private bool CanRemovePreset()
        => SelectedPreset is not null && !SelectedPreset.IsBuiltIn;

    [RelayCommand(CanExecute = nameof(CanAddRecentSelected))]
    private void AddRecentSelected()
    {
        if (!string.IsNullOrWhiteSpace(SelectedRecentFile))
        {
            AddFile(SelectedRecentFile);
        }
    }

    private bool CanAddRecentSelected()
        => !string.IsNullOrWhiteSpace(SelectedRecentFile);

    [RelayCommand]
    private void ClearRecent()
    {
        RecentFiles.Clear();
        AppSettings.SetRecentFiles(RecentFiles);
    }

    [RelayCommand]
    private async Task ImportPresetsAsync()
    {
        var picker = new FileOpenPicker();
        picker.FileTypeFilter.Add(".json");
        InitializePicker(picker);

        var file = await picker.PickSingleFileAsync();
        if (file is null)
        {
            return;
        }

        try
        {
            var json = File.ReadAllText(file.Path);
            var imported = JsonSerializer.Deserialize<List<OutputPreset>>(json) ?? new List<OutputPreset>();
            var added = 0;

            foreach (var preset in imported)
            {
                if (string.IsNullOrWhiteSpace(preset.Name))
                {
                    continue;
                }

                var normalized = new OutputPreset
                {
                    Name = preset.Name.Trim(),
                    OutputFormat = preset.OutputFormat?.Trim() ?? string.Empty,
                    OutputExtension = preset.OutputExtension?.Trim() ?? string.Empty,
                    AdditionalArgs = preset.AdditionalArgs ?? string.Empty,
                    TemplatePath = preset.TemplatePath ?? string.Empty,
                    IsBuiltIn = false
                };

                var existing = customPresets.FirstOrDefault(p => p.Name.Equals(normalized.Name, StringComparison.OrdinalIgnoreCase));
                if (existing is null)
                {
                    customPresets.Add(normalized);
                    added++;
                }
                else
                {
                    existing.OutputFormat = normalized.OutputFormat;
                    existing.OutputExtension = normalized.OutputExtension;
                    existing.AdditionalArgs = normalized.AdditionalArgs;
                    existing.TemplatePath = normalized.TemplatePath;
                }
            }

            AppSettings.SetPresets(customPresets);
            LoadPresets();
            AppendLog($"已导入预设：{added} 项");
        }
        catch (Exception ex)
        {
            AppendLog(ex.Message);
        }
    }

    [RelayCommand]
    private async Task ExportPresetsAsync()
    {
        var picker = new FileSavePicker();
        picker.FileTypeChoices.Add("JSON", new List<string> { ".json" });
        picker.SuggestedFileName = "pandoc-presets";
        InitializePicker(picker);

        var file = await picker.PickSaveFileAsync();
        if (file is null)
        {
            return;
        }

        try
        {
            var json = JsonSerializer.Serialize(customPresets, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(file.Path, json);
            AppendLog($"已导出预设：{customPresets.Count} 项");
        }
        catch (Exception ex)
        {
            AppendLog(ex.Message);
        }
    }

    [RelayCommand(CanExecute = nameof(CanApplyRecentOutputDirectory))]
    private void ApplyRecentOutputDirectory()
    {
        if (!string.IsNullOrWhiteSpace(SelectedRecentOutputDirectory))
        {
            OutputDirectory = SelectedRecentOutputDirectory;
        }
    }

    private bool CanApplyRecentOutputDirectory()
        => !string.IsNullOrWhiteSpace(SelectedRecentOutputDirectory);

    [RelayCommand]
    private void ClearRecentOutputDirectories()
    {
        RecentOutputDirectories.Clear();
        AppSettings.SetRecentOutputDirectories(RecentOutputDirectories);
    }

    [RelayCommand(CanExecute = nameof(CanApplyRecentTemplate))]
    private void ApplyRecentTemplate()
    {
        if (!string.IsNullOrWhiteSpace(SelectedRecentTemplate))
        {
            TemplatePath = SelectedRecentTemplate;
        }
    }

    private bool CanApplyRecentTemplate()
        => !string.IsNullOrWhiteSpace(SelectedRecentTemplate);

    [RelayCommand]
    private void ClearRecentTemplates()
    {
        RecentTemplates.Clear();
        AppSettings.SetRecentTemplates(RecentTemplates);
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
            var pendingItems = Queue.Where(item => item.Status == ConversionStatus.Pending).ToList();
            var parallel = Math.Clamp(MaxParallelism, 1, 8);
            var semaphore = new SemaphoreSlim(parallel, parallel);
            var tasks = pendingItems.Select(item => ProcessItemAsync(item, semaphore, conversionCts.Token)).ToList();
            await Task.WhenAll(tasks);
        }
        catch (OperationCanceledException)
        {
            AppendLog("转换已取消");
            foreach (var item in Queue.Where(item => item.Status == ConversionStatus.Pending))
            {
                item.Status = ConversionStatus.Skipped;
                item.Message = "已取消";
            }
        }
        catch (Exception ex)
        {
            AppendLog(ex.Message);
        }
        finally
        {
            IsBusy = false;
            IsQueuePaused = false;
            lock (pauseLock)
            {
                pauseTcs?.TrySetResult(true);
                pauseTcs = null;
            }
            conversionCts?.Dispose();
            conversionCts = null;
        }
    }

    private bool CanStartConversion()
        => !IsBusy && IsPandocReady && Queue.Count > 0;

    [RelayCommand(CanExecute = nameof(CanPauseQueue))]
    private void PauseQueue()
    {
        if (IsQueuePaused)
        {
            return;
        }

        IsQueuePaused = true;
        lock (pauseLock)
        {
            pauseTcs = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        }

        AppendLog("队列已暂停");
    }

    private bool CanPauseQueue()
        => IsBusy && !IsQueuePaused;

    [RelayCommand(CanExecute = nameof(CanResumeQueue))]
    private void ResumeQueue()
    {
        if (!IsQueuePaused)
        {
            return;
        }

        IsQueuePaused = false;
        lock (pauseLock)
        {
            pauseTcs?.TrySetResult(true);
            pauseTcs = null;
        }

        AppendLog("队列已继续");
    }

    private bool CanResumeQueue()
        => IsBusy && IsQueuePaused;

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
        AddRecentFile(path);
    }

    public void AddFilesFromPaths(IEnumerable<string> paths)
    {
        foreach (var path in paths)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                continue;
            }

            if (Directory.Exists(path))
            {
                foreach (var file in Directory.GetFiles(path))
                {
                    AddFile(file);
                }

                continue;
            }

            if (File.Exists(path))
            {
                AddFile(path);
            }
        }
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

        if (!string.IsNullOrWhiteSpace(TemplatePath))
        {
            args.Add("--template");
            args.Add(TemplatePath);
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

    private Task WaitIfPausedAsync(CancellationToken token)
    {
        Task? waitTask = null;
        lock (pauseLock)
        {
            if (IsQueuePaused)
            {
                pauseTcs ??= new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
                waitTask = pauseTcs.Task;
            }
        }

        if (waitTask is null)
        {
            return Task.CompletedTask;
        }

        return waitTask.WaitAsync(token);
    }

    private static string BuildErrorMessage(string? stderr)
    {
        if (string.IsNullOrWhiteSpace(stderr))
        {
            return "失败";
        }

        var firstLine = stderr
            .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
            .FirstOrDefault();
        return string.IsNullOrWhiteSpace(firstLine) ? "失败" : firstLine.Trim();
    }

    private async Task ProcessItemAsync(ConversionItem item, SemaphoreSlim semaphore, CancellationToken token)
    {
        await semaphore.WaitAsync(token);
        try
        {
            if (item.Status != ConversionStatus.Pending)
            {
                return;
            }

            if (token.IsCancellationRequested)
            {
                item.Status = ConversionStatus.Skipped;
                item.Message = "已取消";
                return;
            }

            await WaitIfPausedAsync(token);

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
            var result = await PandocService.RunAsync(PandocPath, args, token);

            if (result.Succeeded)
            {
                item.Status = ConversionStatus.Succeeded;
                item.Message = "完成";
                AppendLog($"{item.FileName} -> {item.OutputPath}");
            }
            else
            {
                item.Status = ConversionStatus.Failed;
                item.Message = BuildErrorMessage(result.StandardError);
                AppendLog($"失败: {item.FileName}");
            }

            AppendLog(result.StandardError);
        }
        finally
        {
            semaphore.Release();
        }
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
        RetryFailedCommand.NotifyCanExecuteChanged();
        ClearCompletedCommand.NotifyCanExecuteChanged();
        PauseQueueCommand.NotifyCanExecuteChanged();
        ResumeQueueCommand.NotifyCanExecuteChanged();
    }

    private void OnIsPandocReadyChanged(bool value)
    {
        StartConversionCommand.NotifyCanExecuteChanged();
    }

    private void OnIsDownloadingChanged(bool value)
    {
        OnPropertyChanged(nameof(DownloadVisibility));
    }

    [RelayCommand]
    private void ClearLog()
    {
        LogText = string.Empty;
    }

    private void OnPandocPathChanged(string value)
    {
        AppSettings.PandocPath = value;
    }

    private void OnOutputDirectoryChanged(string value)
    {
        AppSettings.OutputDirectory = value;
        UpdateOutputPathsForPending();
        AddRecentOutputDirectory(value);
    }

    private void OnAdditionalArgsChanged(string value)
    {
        AppSettings.AdditionalArgs = value;
    }

    private void OnTemplatePathChanged(string value)
    {
        AppSettings.TemplatePath = value;
        AddRecentTemplate(value);
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

    private void OnSelectedPresetChanged()
    {
        ApplyPresetCommand.NotifyCanExecuteChanged();
        RemovePresetCommand.NotifyCanExecuteChanged();
        if (SelectedPreset is not null)
        {
            PresetName = SelectedPreset.Name;
        }
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

    private void LoadPresets()
    {
        presets.Clear();
        var builtIn = new List<OutputPreset>
        {
            new()
            {
                Name = "PDF",
                OutputFormat = "pdf",
                OutputExtension = "pdf",
                AdditionalArgs = string.Empty,
                IsBuiltIn = true
            },
            new()
            {
                Name = "Word (DOCX)",
                OutputFormat = "docx",
                OutputExtension = "docx",
                AdditionalArgs = string.Empty,
                IsBuiltIn = true
            },
            new()
            {
                Name = "HTML",
                OutputFormat = "html",
                OutputExtension = "html",
                AdditionalArgs = string.Empty,
                IsBuiltIn = true
            },
            new()
            {
                Name = "Markdown",
                OutputFormat = "markdown",
                OutputExtension = "md",
                AdditionalArgs = string.Empty,
                IsBuiltIn = true
            },
            new()
            {
                Name = "PDF (带目录)",
                OutputFormat = "pdf",
                OutputExtension = "pdf",
                AdditionalArgs = "--toc --toc-depth=3",
                IsBuiltIn = true
            },
            new()
            {
                Name = "PDF (高质量)",
                OutputFormat = "pdf",
                OutputExtension = "pdf",
                AdditionalArgs = "--pdf-engine=tectonic",
                IsBuiltIn = true
            },
            new()
            {
                Name = "PDF (低质量)",
                OutputFormat = "pdf",
                OutputExtension = "pdf",
                AdditionalArgs = "--pdf-engine=pdflatex",
                IsBuiltIn = true
            }
        };

        foreach (var preset in builtIn)
        {
            presets.Add(preset);
        }

        customPresets = AppSettings.Presets.ToList();
        foreach (var preset in customPresets)
        {
            presets.Add(preset);
        }

        SelectedPreset = presets.FirstOrDefault(p => p.Name.Equals(PresetName, StringComparison.OrdinalIgnoreCase))
            ?? presets.FirstOrDefault();
    }

    private void LoadRecentFiles()
    {
        recentFiles.Clear();
        foreach (var file in AppSettings.RecentFiles)
        {
            recentFiles.Add(file);
        }
    }

    private void LoadRecentOutputDirectories()
    {
        recentOutputDirectories.Clear();
        foreach (var directory in AppSettings.RecentOutputDirectories)
        {
            recentOutputDirectories.Add(directory);
        }
    }

    private void LoadRecentTemplates()
    {
        recentTemplates.Clear();
        foreach (var template in AppSettings.RecentTemplates)
        {
            recentTemplates.Add(template);
        }
    }

    private void AddRecentFile(string path)
    {
        var existing = recentFiles.FirstOrDefault(item => item.Equals(path, StringComparison.OrdinalIgnoreCase));
        if (existing is not null)
        {
            recentFiles.Remove(existing);
        }

        recentFiles.Insert(0, path);
        while (recentFiles.Count > 10)
        {
            recentFiles.RemoveAt(recentFiles.Count - 1);
        }

        AppSettings.SetRecentFiles(recentFiles);
    }

    private void AddRecentOutputDirectory(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return;
        }

        var existing = recentOutputDirectories.FirstOrDefault(item => item.Equals(path, StringComparison.OrdinalIgnoreCase));
        if (existing is not null)
        {
            recentOutputDirectories.Remove(existing);
        }

        recentOutputDirectories.Insert(0, path);
        while (recentOutputDirectories.Count > 10)
        {
            recentOutputDirectories.RemoveAt(recentOutputDirectories.Count - 1);
        }

        AppSettings.SetRecentOutputDirectories(recentOutputDirectories);
    }

    private void AddRecentTemplate(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return;
        }

        var existing = recentTemplates.FirstOrDefault(item => item.Equals(path, StringComparison.OrdinalIgnoreCase));
        if (existing is not null)
        {
            recentTemplates.Remove(existing);
        }

        recentTemplates.Insert(0, path);
        while (recentTemplates.Count > 10)
        {
            recentTemplates.RemoveAt(recentTemplates.Count - 1);
        }

        AppSettings.SetRecentTemplates(recentTemplates);
    }
}
