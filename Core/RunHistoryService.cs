using System.IO;
using System.Text.Json;
using SmartMacroAI.Models;

namespace SmartMacroAI.Core;

/// <summary>
/// Manages macro run history persistence (JSON per macro, last 100 runs each).
/// Created by Phạm Duy – Giải pháp tự động hóa thông minh.
/// </summary>
public sealed class RunHistoryService
{
    private const int MaxRecords = 100;
    private readonly string _historyFolder;

    public RunHistoryService()
    {
        _historyFolder = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory, "logs", "history");
        Directory.CreateDirectory(_historyFolder);
    }

    public void Save(MacroRunRecord record)
    {
        string file = Path.Combine(_historyFolder, $"{SanitizeFileName(record.MacroName)}.json");

        var list = Load(record.MacroName);
        list.Insert(0, record);

        if (list.Count > MaxRecords)
            list = list.Take(MaxRecords).ToList();

        try
        {
            File.WriteAllText(file,
                JsonSerializer.Serialize(list, new JsonSerializerOptions { WriteIndented = true }));
        }
        catch
        {
            // Non-critical — don't crash macro execution for log failures
        }
    }

    public List<MacroRunRecord> Load(string macroName)
    {
        string file = Path.Combine(_historyFolder, $"{SanitizeFileName(macroName)}.json");
        if (!File.Exists(file))
            return new List<MacroRunRecord>();

        try
        {
            return JsonSerializer.Deserialize<List<MacroRunRecord>>(
                File.ReadAllText(file)) ?? new List<MacroRunRecord>();
        }
        catch
        {
            return new List<MacroRunRecord>();
        }
    }

    public List<MacroRunRecord> LoadAll()
    {
        var results = new List<MacroRunRecord>();

        try
        {
            foreach (string file in Directory.GetFiles(_historyFolder, "*.json"))
            {
                try
                {
                    var records = JsonSerializer.Deserialize<List<MacroRunRecord>>(
                        File.ReadAllText(file));
                    if (records != null)
                        results.AddRange(records);
                }
                catch
                {
                    // Skip corrupted files
                }
            }
        }
        catch
        {
            // Folder may not exist
        }

        return results
            .OrderByDescending(r => r.StartTime)
            .Take(200)
            .ToList();
    }

    public void Clear(string? macroName = null)
    {
        if (string.IsNullOrWhiteSpace(macroName))
        {
            // Clear all history
            try
            {
                foreach (string file in Directory.GetFiles(_historyFolder, "*.json"))
                {
                    try { File.Delete(file); } catch { }
                }
            }
            catch { }
        }
        else
        {
            // Clear specific macro history
            string file = Path.Combine(_historyFolder, $"{SanitizeFileName(macroName)}.json");
            try { File.Delete(file); } catch { }
        }
    }

    private static string SanitizeFileName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            return "unknown";

        char[] invalidChars = Path.GetInvalidFileNameChars();
        return string.Concat(name.Split(invalidChars));
    }
}
