using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using SmartMacroAI.Models;

namespace SmartMacroAI.Core;

/// <summary>
/// Handles serialization and deserialization of <see cref="MacroScript"/> objects
/// to/from JSON files. Uses System.Text.Json with polymorphic type support
/// so that all concrete <see cref="MacroAction"/> subtypes survive round-trips.
/// </summary>
public static class ScriptManager
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Converters = { new JsonStringEnumConverter() },
    };

    public static string DefaultScriptsFolder
    {
        get
        {
            string folder = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory, "Scripts");
            Directory.CreateDirectory(folder);
            return folder;
        }
    }

    /// <summary>
    /// Serializes a <see cref="MacroScript"/> and writes it to <paramref name="filePath"/>.
    /// </summary>
    public static async Task SaveAsync(MacroScript script, string filePath)
    {
        ArgumentNullException.ThrowIfNull(script);
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);

        script.ModifiedAt = DateTime.UtcNow;

        string directory = Path.GetDirectoryName(filePath)!;
        if (!string.IsNullOrEmpty(directory))
            Directory.CreateDirectory(directory);

        string json = JsonSerializer.Serialize(script, JsonOptions);
        await File.WriteAllTextAsync(filePath, json);
    }

    /// <summary>
    /// Synchronous overload for quick saves (e.g. auto-save on UI close).
    /// </summary>
    public static void Save(MacroScript script, string filePath)
    {
        ArgumentNullException.ThrowIfNull(script);
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);

        script.ModifiedAt = DateTime.UtcNow;

        string directory = Path.GetDirectoryName(filePath)!;
        if (!string.IsNullOrEmpty(directory))
            Directory.CreateDirectory(directory);

        string json = JsonSerializer.Serialize(script, JsonOptions);
        File.WriteAllText(filePath, json);
    }

    /// <summary>
    /// Reads a JSON file and deserializes it into a <see cref="MacroScript"/>.
    /// Returns null if the file does not exist or deserialization fails.
    /// </summary>
    public static async Task<MacroScript?> LoadAsync(string filePath)
    {
        if (!File.Exists(filePath))
            return null;

        string json = await File.ReadAllTextAsync(filePath);
        return JsonSerializer.Deserialize<MacroScript>(json, JsonOptions);
    }

    /// <summary>
    /// Synchronous overload for loading.
    /// </summary>
    public static MacroScript? Load(string filePath)
    {
        if (!File.Exists(filePath))
            return null;

        string json = File.ReadAllText(filePath);
        return JsonSerializer.Deserialize<MacroScript>(json, JsonOptions);
    }

    /// <summary>
    /// Returns all .json script files found in the default scripts folder.
    /// </summary>
    public static IEnumerable<string> EnumerateSavedScripts()
    {
        return Directory.Exists(DefaultScriptsFolder)
            ? Directory.EnumerateFiles(DefaultScriptsFolder, "*.json")
            : Enumerable.Empty<string>();
    }

    /// <summary>
    /// Builds a suggested file path for a script in the default folder,
    /// sanitizing the script name for use as a filename.
    /// </summary>
    public static string GetDefaultFilePath(MacroScript script)
    {
        string safeName = SanitizeFileStem(script.Name);
        return Path.Combine(DefaultScriptsFolder, safeName + ".json");
    }

    /// <summary>Turns a display name into a safe filename stem (no extension).</summary>
    public static string SanitizeFileStem(string name)
    {
        string safe = string.Join("_",
            (name ?? "").Split(Path.GetInvalidFileNameChars(), StringSplitOptions.RemoveEmptyEntries)).Trim();
        return string.IsNullOrWhiteSpace(safe) ? "Macro" : safe;
    }
}
