using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
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
        Queue.CollectionChanged += OnQueueCollectionChanged;
    }

    public ObservableCollection<ConversionItem> Queue { get; } = new();

    public ObservableCollection<string> InputFormats { get; } = new();

    public ObservableCollection<string> OutputFormats { get; } = new();

    private bool isBusy;
    private bool isPandocReady;
    private bool isQueuePaused;
    private string pandocPath = string.Empty;
    private string pandocStatus = "未检测";
    private bool autoStartOnDrop;
    private string selectedInputFormat = "Auto";
    private string selectedOutputFormat = "docx";
    private string outputDirectory = string.Empty;
    private string outputExtension = "docx";
    private string additionalArgs = string.Empty;
    private string templatePath = string.Empty;
    private string presetName = string.Empty;
    private string logText = string.Empty;
    private string workflowStatusText = string.Empty;
    private string selectedFormatSetting = string.Empty;
    private string formatSettingOutputExtension = string.Empty;
    private string formatSettingAdditionalArgs = string.Empty;
    private string formatSettingTemplatePath = string.Empty;
    private ConversionItem? selectedItem;
    private bool isDownloading;
    private double downloadProgress;
    private TaskCompletionSource<bool>? pauseTcs;
    private readonly object pauseLock = new();
    private int maxParallelism = 1;
    private string? selectedRecentOutputDirectory;
    private string? selectedRecentTemplate;
    private int queueTotal;
    private int queuePending;
    private int queueRunning;
    private int queueSucceeded;
    private int queueFailed;
    private int queueSkipped;
    private int queueCompleted;
    private double queueProgress;
    private OutputPreset? selectedPreset;
    private string? selectedRecentFile;
    private readonly ObservableCollection<OutputPreset> presets = new();
    private readonly ObservableCollection<string> recentFiles = new();
    private readonly ObservableCollection<string> recentOutputDirectories = new();
    private readonly ObservableCollection<string> recentOutputFormats = new();
    private readonly ObservableCollection<string> recentTemplates = new();
    private readonly ObservableCollection<int> parallelismOptions = new() { 1, 2, 3, 4, 5, 6, 7, 8 };
    private readonly Dictionary<string, OutputFormatSettings> formatSettings = new(StringComparer.OrdinalIgnoreCase);
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

    public bool AutoStartOnDrop
    {
        get => autoStartOnDrop;
        set
        {
            if (SetProperty(ref autoStartOnDrop, value))
            {
                AppSettings.AutoStartOnDrop = value;
            }
        }
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

    public string WorkflowStatusText
    {
        get => workflowStatusText;
        private set => SetProperty(ref workflowStatusText, value);
    }

    public string SelectedFormatSetting
    {
        get => selectedFormatSetting;
        set
        {
            if (SetProperty(ref selectedFormatSetting, value))
            {
                LoadFormatSettingFields(value);
                SaveFormatSettingCommand.NotifyCanExecuteChanged();
                ClearFormatSettingCommand.NotifyCanExecuteChanged();
            }
        }
    }

    public string FormatSettingOutputExtension
    {
        get => formatSettingOutputExtension;
        set
        {
            if (SetProperty(ref formatSettingOutputExtension, value))
            {
                SaveFormatSettingCommand.NotifyCanExecuteChanged();
            }
        }
    }

    public string FormatSettingAdditionalArgs
    {
        get => formatSettingAdditionalArgs;
        set
        {
            if (SetProperty(ref formatSettingAdditionalArgs, value))
            {
                SaveFormatSettingCommand.NotifyCanExecuteChanged();
            }
        }
    }

    public string FormatSettingTemplatePath
    {
        get => formatSettingTemplatePath;
        set
        {
            if (SetProperty(ref formatSettingTemplatePath, value))
            {
                SaveFormatSettingCommand.NotifyCanExecuteChanged();
            }
        }
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

    public int QueueTotal
    {
        get => queueTotal;
        private set => SetProperty(ref queueTotal, value);
    }

    public int QueuePending
    {
        get => queuePending;
        private set => SetProperty(ref queuePending, value);
    }

    public int QueueRunning
    {
        get => queueRunning;
        private set => SetProperty(ref queueRunning, value);
    }

    public int QueueSucceeded
    {
        get => queueSucceeded;
        private set => SetProperty(ref queueSucceeded, value);
    }

    public int QueueFailed
    {
        get => queueFailed;
        private set => SetProperty(ref queueFailed, value);
    }

    public int QueueSkipped
    {
        get => queueSkipped;
        private set => SetProperty(ref queueSkipped, value);
    }

    public int QueueCompleted
    {
        get => queueCompleted;
        private set => SetProperty(ref queueCompleted, value);
    }

    public double QueueProgress
    {
        get => queueProgress;
        private set => SetProperty(ref queueProgress, value);
    }

    public string QueueSummary => $"总计 {QueueTotal} · 进行中 {QueueRunning} · 待处理 {QueuePending}";

    public string QueueResultSummary => $"完成 {QueueSucceeded} · 失败 {QueueFailed} · 跳过 {QueueSkipped}";

    public string QueueProgressText => $"进度 {QueueCompleted}/{QueueTotal}";

    public Visibility BusyVisibility => IsBusy ? Visibility.Visible : Visibility.Collapsed;

    public Visibility DownloadVisibility => IsDownloading ? Visibility.Visible : Visibility.Collapsed;

    public int QueueCount => Queue.Count;

    public ObservableCollection<string> RecentOutputFormats => recentOutputFormats;

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
        AutoStartOnDrop = AppSettings.AutoStartOnDrop;
        LoadPresets();
        LoadFormatSettings();
        LoadRecentFiles();
        LoadRecentOutputDirectories();
        LoadRecentOutputFormats();
        LoadRecentTemplates();
        UpdateQueueStats();
        if (!AppSettings.HasLaunchedBefore)
        {
            AppSettings.HasLaunchedBefore = true;
        }

        await RefreshPandocAsync();
        UpdateWorkflowStatusText();
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
    private async Task BrowseFormatTemplateAsync()
    {
        var picker = new FileOpenPicker();
        picker.FileTypeFilter.Add("*");
        InitializePicker(picker);

        var file = await picker.PickSingleFileAsync();
        if (file is not null)
        {
            FormatSettingTemplatePath = file.Path;
        }
    }

    [RelayCommand(CanExecute = nameof(CanSaveFormatSetting))]
    private void SaveFormatSetting()
    {
        var format = SelectedFormatSetting?.Trim();
        if (string.IsNullOrWhiteSpace(format))
        {
            return;
        }

        var outputExtension = NormalizeSettingValue(FormatSettingOutputExtension);
        var additionalArgs = NormalizeSettingValue(FormatSettingAdditionalArgs);
        var templatePath = NormalizeSettingValue(FormatSettingTemplatePath);
        var hasOverrides = !string.IsNullOrWhiteSpace(outputExtension) ||
            !string.IsNullOrWhiteSpace(additionalArgs) ||
            !string.IsNullOrWhiteSpace(templatePath);

        if (!hasOverrides)
        {
            formatSettings.Remove(format);
        }
        else
        {
            formatSettings[format] = new OutputFormatSettings
            {
                Format = format,
                OutputExtension = outputExtension,
                AdditionalArgs = additionalArgs,
                TemplatePath = templatePath
            };
        }

        AppSettings.SetOutputFormatSettings(formatSettings.Values);
        if (string.Equals(format, SelectedOutputFormat, StringComparison.OrdinalIgnoreCase))
        {
            UpdateOutputPathsForPending();
        }

        ClearFormatSettingCommand.NotifyCanExecuteChanged();
        SaveFormatSettingCommand.NotifyCanExecuteChanged();
    }

    private bool CanSaveFormatSetting()
        => !string.IsNullOrWhiteSpace(SelectedFormatSetting);

    [RelayCommand(CanExecute = nameof(CanClearFormatSetting))]
    private void ClearFormatSetting()
    {
        var format = SelectedFormatSetting?.Trim();
        if (string.IsNullOrWhiteSpace(format))
        {
            return;
        }

        formatSettings.Remove(format);
        AppSettings.SetOutputFormatSettings(formatSettings.Values);
        FormatSettingOutputExtension = string.Empty;
        FormatSettingAdditionalArgs = string.Empty;
        FormatSettingTemplatePath = string.Empty;
        if (string.Equals(format, SelectedOutputFormat, StringComparison.OrdinalIgnoreCase))
        {
            UpdateOutputPathsForPending();
        }

        ClearFormatSettingCommand.NotifyCanExecuteChanged();
        SaveFormatSettingCommand.NotifyCanExecuteChanged();
    }

    private bool CanClearFormatSetting()
    {
        var format = SelectedFormatSetting?.Trim();
        return !string.IsNullOrWhiteSpace(format) && formatSettings.ContainsKey(format);
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

    [RelayCommand]
    private void SetOutputFormat(string? format)
    {
        if (string.IsNullOrWhiteSpace(format))
        {
            return;
        }

        SelectedOutputFormat = format;
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

    private void OnQueueCollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        if (e.NewItems is not null)
        {
            foreach (var item in e.NewItems.OfType<ConversionItem>())
            {
                item.PropertyChanged += OnQueueItemPropertyChanged;
            }
        }

        if (e.OldItems is not null)
        {
            foreach (var item in e.OldItems.OfType<ConversionItem>())
            {
                item.PropertyChanged -= OnQueueItemPropertyChanged;
            }
        }

        StartConversionCommand.NotifyCanExecuteChanged();
        OnPropertyChanged(nameof(QueueCount));
        RetryFailedCommand.NotifyCanExecuteChanged();
        ClearCompletedCommand.NotifyCanExecuteChanged();
        UpdateQueueStats();
        UpdateWorkflowStatusText();
    }

    private void OnQueueItemPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(ConversionItem.Status))
        {
            UpdateQueueStats();
            RetryFailedCommand.NotifyCanExecuteChanged();
            ClearCompletedCommand.NotifyCanExecuteChanged();
        }
    }

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

        var extension = GetEffectiveOutputExtension();
        if (string.IsNullOrWhiteSpace(extension))
        {
            extension = SelectedOutputFormat;
        }

        extension = extension.TrimStart('.');
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

        var formatSetting = GetFormatSetting(SelectedOutputFormat);
        var templatePath = NormalizeSettingValue(formatSetting?.TemplatePath) ?? TemplatePath;
        if (!string.IsNullOrWhiteSpace(templatePath))
        {
            args.Add("--template");
            args.Add(templatePath);
        }

        args.AddRange(CommandLineTokenizer.Split(AdditionalArgs));
        var formatArgs = NormalizeSettingValue(formatSetting?.AdditionalArgs);
        if (!string.IsNullOrWhiteSpace(formatArgs))
        {
            args.AddRange(CommandLineTokenizer.Split(formatArgs));
        }
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

        if (OutputFormats.Count > 0 &&
            (string.IsNullOrWhiteSpace(SelectedFormatSetting) ||
            !OutputFormats.Any(f => f.Equals(SelectedFormatSetting, StringComparison.OrdinalIgnoreCase))))
        {
            SelectedFormatSetting = SelectedOutputFormat;
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
        UpdateWorkflowStatusText();
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
        UpdateWorkflowStatusText();
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
        AddRecentOutputFormat(value);
        UpdateOutputPathsForPending();
        UpdateWorkflowStatusText();
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

    private void UpdateQueueStats()
    {
        var total = Queue.Count;
        var pending = Queue.Count(item => item.Status == ConversionStatus.Pending);
        var running = Queue.Count(item => item.Status == ConversionStatus.Running);
        var succeeded = Queue.Count(item => item.Status == ConversionStatus.Succeeded);
        var failed = Queue.Count(item => item.Status == ConversionStatus.Failed);
        var skipped = Queue.Count(item => item.Status == ConversionStatus.Skipped);
        var completed = succeeded + failed + skipped;

        QueueTotal = total;
        QueuePending = pending;
        QueueRunning = running;
        QueueSucceeded = succeeded;
        QueueFailed = failed;
        QueueSkipped = skipped;
        QueueCompleted = completed;
        QueueProgress = total == 0 ? 0 : (double)completed / total;
        OnPropertyChanged(nameof(QueueSummary));
        OnPropertyChanged(nameof(QueueResultSummary));
        OnPropertyChanged(nameof(QueueProgressText));
    }

    private void UpdateWorkflowStatusText()
    {
        var importText = Queue.Count == 0 ? "未导入文件" : $"已导入 {Queue.Count} 个文件";
        var formatText = string.IsNullOrWhiteSpace(SelectedOutputFormat)
            ? "待选择输出格式"
            : $"输出格式 {SelectedOutputFormat}";
        var outputText = string.IsNullOrWhiteSpace(OutputDirectory) ? "输出目录未设置" : "输出目录已设置";
        var pandocText = IsPandocReady ? "Pandoc 已就绪" : "Pandoc 未就绪";
        WorkflowStatusText = $"{importText} · {formatText} · {outputText} · {pandocText}";
    }

    private void LoadFormatSettings()
    {
        formatSettings.Clear();
        foreach (var setting in AppSettings.OutputFormatSettings)
        {
            if (string.IsNullOrWhiteSpace(setting.Format))
            {
                continue;
            }

            formatSettings[setting.Format] = setting;
        }

        if (string.IsNullOrWhiteSpace(SelectedFormatSetting))
        {
            SelectedFormatSetting = SelectedOutputFormat;
        }
        else
        {
            LoadFormatSettingFields(SelectedFormatSetting);
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
            },
            new()
            {
                Name = "HTML + CSS",
                OutputFormat = "html",
                OutputExtension = "html",
                AdditionalArgs = "--css=style.css",
                IsBuiltIn = true
            },
            new()
            {
                Name = "Word 模板",
                OutputFormat = "docx",
                OutputExtension = "docx",
                AdditionalArgs = "--reference-doc=template.docx",
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

    private void LoadRecentOutputFormats()
    {
        recentOutputFormats.Clear();
        foreach (var format in AppSettings.RecentOutputFormats)
        {
            if (!string.IsNullOrWhiteSpace(format))
            {
                recentOutputFormats.Add(format);
            }
        }

        if (recentOutputFormats.Count == 0 && !string.IsNullOrWhiteSpace(SelectedOutputFormat))
        {
            recentOutputFormats.Add(SelectedOutputFormat);
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

    private void LoadFormatSettingFields(string? format)
    {
        if (string.IsNullOrWhiteSpace(format))
        {
            FormatSettingOutputExtension = string.Empty;
            FormatSettingAdditionalArgs = string.Empty;
            FormatSettingTemplatePath = string.Empty;
            return;
        }

        if (formatSettings.TryGetValue(format.Trim(), out var setting))
        {
            FormatSettingOutputExtension = setting.OutputExtension ?? string.Empty;
            FormatSettingAdditionalArgs = setting.AdditionalArgs ?? string.Empty;
            FormatSettingTemplatePath = setting.TemplatePath ?? string.Empty;
            return;
        }

        FormatSettingOutputExtension = string.Empty;
        FormatSettingAdditionalArgs = string.Empty;
        FormatSettingTemplatePath = string.Empty;
    }

    private OutputFormatSettings? GetFormatSetting(string? format)
    {
        if (string.IsNullOrWhiteSpace(format))
        {
            return null;
        }

        return formatSettings.TryGetValue(format.Trim(), out var setting) ? setting : null;
    }

    private static string? NormalizeSettingValue(string? value)
    {
        var trimmed = value?.Trim();
        return string.IsNullOrWhiteSpace(trimmed) ? null : trimmed;
    }

    private string GetEffectiveOutputExtension()
    {
        var formatSetting = GetFormatSetting(SelectedOutputFormat);
        return NormalizeSettingValue(formatSetting?.OutputExtension) ?? OutputExtension;
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

    private void AddRecentOutputFormat(string format)
    {
        if (string.IsNullOrWhiteSpace(format))
        {
            return;
        }

        var normalized = format.Trim();
        var existing = recentOutputFormats.FirstOrDefault(item => item.Equals(normalized, StringComparison.OrdinalIgnoreCase));
        if (existing is not null)
        {
            recentOutputFormats.Remove(existing);
        }

        recentOutputFormats.Insert(0, normalized);
        while (recentOutputFormats.Count > 6)
        {
            recentOutputFormats.RemoveAt(recentOutputFormats.Count - 1);
        }

        AppSettings.SetRecentOutputFormats(recentOutputFormats);
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
