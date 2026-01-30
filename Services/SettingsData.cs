using PandocGUI.Models;

namespace PandocGUI.Services;

public sealed class SettingsData
{
    public bool HasLaunchedBefore { get; set; }

    public string? PandocPath { get; set; }

    public string? OutputDirectory { get; set; }

    public string? AdditionalArgs { get; set; }

    public string? SelectedOutputFormat { get; set; }

    public string? SelectedInputFormat { get; set; }

    public string? OutputExtension { get; set; }

    public string? TemplatePath { get; set; }

    public List<OutputPreset> Presets { get; set; } = new();

    public List<string> RecentFiles { get; set; } = new();
}
