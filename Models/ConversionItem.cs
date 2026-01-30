using System.IO;

namespace PandocGUI.Models;

public partial class ConversionItem : ObservableObject
{
    public ConversionItem(string inputPath)
    {
        InputPath = inputPath;
        FileName = Path.GetFileName(inputPath);
    }

    public string InputPath { get; }

    public string FileName { get; }

    private string outputPath = string.Empty;
    private ConversionStatus status = ConversionStatus.Pending;
    private string message = string.Empty;

    public string OutputPath
    {
        get => outputPath;
        set => SetProperty(ref outputPath, value);
    }

    public ConversionStatus Status
    {
        get => status;
        set => SetProperty(ref status, value);
    }

    public string Message
    {
        get => message;
        set => SetProperty(ref message, value);
    }
}
