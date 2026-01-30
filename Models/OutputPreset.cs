namespace PandocGUI.Models;

public sealed class OutputPreset
{
    public string Name { get; set; } = string.Empty;

    public string OutputFormat { get; set; } = string.Empty;

    public string OutputExtension { get; set; } = string.Empty;

    public string AdditionalArgs { get; set; } = string.Empty;

    public string TemplatePath { get; set; } = string.Empty;

    public bool IsBuiltIn { get; set; }
}
