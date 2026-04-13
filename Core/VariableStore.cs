// Created by Phạm Duy – Giải pháp tự động hóa thông minh.

using System.Collections.Concurrent;

namespace SmartMacroAI.Core;

/// <summary>
/// Thread-safe string store for runtime macro variables (used with <c>{{name}}</c> interpolation).
/// </summary>
public sealed class VariableStore
{
    private readonly ConcurrentDictionary<string, (string Value, string Source)> _data =
        new(StringComparer.OrdinalIgnoreCase);

    private static string NormalizeName(string name)
    {
        name = name.Trim();
        if (name.StartsWith("{{", StringComparison.Ordinal) && name.EndsWith("}}", StringComparison.Ordinal))
            name = name.Substring(2, name.Length - 4).Trim();
        return name;
    }

    /// <summary>Sets a variable; <paramref name="source"/> is one of Manual, OCR, Clipboard.</summary>
    public void Set(string name, string value, string source = "Manual")
    {
        string key = NormalizeName(name);
        if (key.Length == 0)
            return;
        _data[key] = (value ?? string.Empty, string.IsNullOrWhiteSpace(source) ? "Manual" : source.Trim());
    }

    public bool TryGet(string name, out string value)
    {
        if (_data.TryGetValue(NormalizeName(name), out var tuple))
        {
            value = tuple.Value;
            return true;
        }

        value = string.Empty;
        return false;
    }

    public string Get(string name, string defaultValue = "") =>
        TryGet(name, out string v) ? v : defaultValue;

    public bool Exists(string name) => _data.ContainsKey(NormalizeName(name));

    public void Clear() => _data.Clear();

    /// <summary>Removes one variable; if <paramref name="name"/> is null/empty, clears all.</summary>
    public void Remove(string? name)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            Clear();
            return;
        }

        _data.TryRemove(NormalizeName(name), out _);
    }

    /// <summary>Snapshot for UI / interpolation (copy).</summary>
    public IReadOnlyDictionary<string, (string Value, string Source)> Snapshot() =>
        new Dictionary<string, (string, string)>(_data, StringComparer.OrdinalIgnoreCase);
}
