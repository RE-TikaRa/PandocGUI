namespace PandocGUI.Models;

public sealed class OutputFormatSettings
{
    public string Format { get; set; } = string.Empty;

    public string? OutputExtension { get; set; }

    public string? AdditionalArgs { get; set; }

    public string? TemplatePath { get; set; }
}
