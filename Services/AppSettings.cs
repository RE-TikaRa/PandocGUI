using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using PandocGUI.Models;

namespace PandocGUI.Services;

public static class AppSettings
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true
    };

    private static SettingsData data = Load();

    public static string SettingsFilePath => Path.Combine(DataDirectoryPath, "settings.json");

    private static string DataDirectoryPath => Path.Combine(AppContext.BaseDirectory, "data");

    public static bool HasLaunchedBefore
    {
        get => data.HasLaunchedBefore;
        set
        {
            if (data.HasLaunchedBefore == value)
            {
                return;
            }

            data.HasLaunchedBefore = value;
            Save();
        }
    }

    public static string? PandocPath
    {
        get => data.PandocPath;
        set => SetString(data.PandocPath, value, v => data.PandocPath = v);
    }

    public static string? OutputDirectory
    {
        get => data.OutputDirectory;
        set => SetString(data.OutputDirectory, value, v => data.OutputDirectory = v);
    }

    public static string? AdditionalArgs
    {
        get => data.AdditionalArgs;
        set => SetString(data.AdditionalArgs, value, v => data.AdditionalArgs = v);
    }

    public static string? SelectedOutputFormat
    {
        get => data.SelectedOutputFormat;
        set => SetString(data.SelectedOutputFormat, value, v => data.SelectedOutputFormat = v);
    }

    public static string? SelectedInputFormat
    {
        get => data.SelectedInputFormat;
        set => SetString(data.SelectedInputFormat, value, v => data.SelectedInputFormat = v);
    }

    public static string? OutputExtension
    {
        get => data.OutputExtension;
        set => SetString(data.OutputExtension, value, v => data.OutputExtension = v);
    }

    public static string? TemplatePath
    {
        get => data.TemplatePath;
        set => SetString(data.TemplatePath, value, v => data.TemplatePath = v);
    }

    public static int MaxParallelism
    {
        get => data.MaxParallelism <= 0 ? 1 : data.MaxParallelism;
        set
        {
            var normalized = value <= 0 ? 1 : value;
            if (data.MaxParallelism == normalized)
            {
                return;
            }

            data.MaxParallelism = normalized;
            Save();
        }
    }

    public static IReadOnlyList<OutputPreset> Presets => data.Presets;

    public static IReadOnlyList<string> RecentFiles => data.RecentFiles;

    public static IReadOnlyList<string> RecentOutputDirectories => data.RecentOutputDirectories;

    public static IReadOnlyList<string> RecentTemplates => data.RecentTemplates;

    public static void SetPresets(IEnumerable<OutputPreset> presets)
    {
        data.Presets = presets.ToList();
        Save();
    }

    public static void SetRecentFiles(IEnumerable<string> files)
    {
        data.RecentFiles = files.ToList();
        Save();
    }

    public static void SetRecentOutputDirectories(IEnumerable<string> directories)
    {
        data.RecentOutputDirectories = directories.ToList();
        Save();
    }

    public static void SetRecentTemplates(IEnumerable<string> templates)
    {
        data.RecentTemplates = templates.ToList();
        Save();
    }

    private static void SetString(string? current, string? value, Action<string?> apply)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            value = null;
        }

        if (string.Equals(current, value, StringComparison.Ordinal))
        {
            return;
        }

        apply(value);
        Save();
    }

    private static SettingsData Load()
    {
        try
        {
            var path = SettingsFilePath;
            if (!File.Exists(path))
            {
                return new SettingsData();
            }

            var json = File.ReadAllText(path);
            var loaded = JsonSerializer.Deserialize<SettingsData>(json) ?? new SettingsData();
            loaded.Presets ??= new List<OutputPreset>();
            loaded.RecentFiles ??= new List<string>();
            loaded.RecentOutputDirectories ??= new List<string>();
            loaded.RecentTemplates ??= new List<string>();
            return loaded;
        }
        catch
        {
            return new SettingsData();
        }
    }

    private static void Save()
    {
        Directory.CreateDirectory(DataDirectoryPath);
        var json = JsonSerializer.Serialize(data, JsonOptions);
        File.WriteAllText(SettingsFilePath, json);
    }
}
