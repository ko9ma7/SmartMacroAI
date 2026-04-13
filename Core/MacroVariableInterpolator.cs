using System.Text.RegularExpressions;

namespace SmartMacroAI.Core;

/// <summary>
/// Expands <c>${VariableName}</c> and <c>{{variable_name}}</c> placeholders using runtime maps.
/// Created by Phạm Duy – Giải pháp tự động hóa thông minh.
/// </summary>
public static class MacroVariableInterpolator
{
    private static readonly Regex Placeholder = new(
        @"\$\{([^}]+)\}",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex DoubleCurly = new(
        @"\{\{([^}]+)\}\}",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    public static string Expand(string input, IReadOnlyDictionary<string, string> vars)
    {
        if (string.IsNullOrEmpty(input) || vars.Count == 0)
            return input;

        return Placeholder.Replace(input, m =>
        {
            string key = m.Groups[1].Value.Trim();
            return vars.TryGetValue(key, out string? v) ? v : m.Value;
        });
    }

    /// <summary>
    /// Replaces <c>{{name}}</c> tokens. Missing keys invoke <paramref name="onMissing"/> and become empty string.
    /// </summary>
    public static string ExpandDoubleCurly(
        string input,
        IReadOnlyDictionary<string, string> values,
        Action<string>? onMissing = null)
    {
        if (string.IsNullOrEmpty(input))
            return input;

        return DoubleCurly.Replace(input, m =>
        {
            string key = m.Groups[1].Value.Trim();
            if (values.TryGetValue(key, out string? v))
                return v ?? string.Empty;
            onMissing?.Invoke(key);
            return string.Empty;
        });
    }
}
