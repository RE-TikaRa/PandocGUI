using System.Collections.Generic;
using System.Text.Json.Serialization;
using PandocGUI.Models;

namespace PandocGUI.Services;

[JsonSourceGenerationOptions(WriteIndented = true)]
[JsonSerializable(typeof(SettingsData), TypeInfoPropertyName = "SettingsData")]
[JsonSerializable(typeof(OutputPreset), TypeInfoPropertyName = "OutputPreset")]
[JsonSerializable(typeof(OutputFormatSettings), TypeInfoPropertyName = "OutputFormatSettings")]
[JsonSerializable(typeof(List<OutputPreset>), TypeInfoPropertyName = "OutputPresetList")]
[JsonSerializable(typeof(List<OutputFormatSettings>), TypeInfoPropertyName = "OutputFormatSettingsList")]
[JsonSerializable(typeof(List<string>), TypeInfoPropertyName = "StringList")]
internal partial class AppJsonContext : JsonSerializerContext
{
}
