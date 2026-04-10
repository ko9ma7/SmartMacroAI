using System.IO;
using System.Text.Json;

namespace SmartMacroAI.Core;

/// <summary>
/// Persists global hotkey configuration to a JSON file next to the executable.
/// Default hotkeys: Ctrl+F5 (toggle app), Ctrl+F7 (toggle target).
/// </summary>
public class HotkeySettings
{
    public const int MOD_ALT     = 0x0001;
    public const int MOD_CONTROL = 0x0002;
    public const int MOD_SHIFT   = 0x0004;

    public int ToggleAppModifier    { get; set; } = MOD_CONTROL;
    public int ToggleAppKey         { get; set; } = 0x74; // VK_F5
    public int ToggleTargetModifier { get; set; } = MOD_CONTROL;
    public int ToggleTargetKey      { get; set; } = 0x76; // VK_F7

    private static readonly string SettingsPath = Path.Combine(
        AppDomain.CurrentDomain.BaseDirectory, "hotkey_settings.json");

    public static HotkeySettings Load()
    {
        try
        {
            if (File.Exists(SettingsPath))
            {
                string json = File.ReadAllText(SettingsPath);
                return JsonSerializer.Deserialize<HotkeySettings>(json) ?? new();
            }
        }
        catch { }
        return new HotkeySettings();
    }

    public void Save()
    {
        try
        {
            string json = JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(SettingsPath, json);
        }
        catch { }
    }

    public static string ModifierToString(int mod) => mod switch
    {
        MOD_ALT => "Alt",
        MOD_CONTROL => "Ctrl",
        MOD_SHIFT => "Shift",
        MOD_CONTROL | MOD_ALT => "Ctrl+Alt",
        MOD_CONTROL | MOD_SHIFT => "Ctrl+Shift",
        MOD_ALT | MOD_SHIFT => "Alt+Shift",
        MOD_CONTROL | MOD_ALT | MOD_SHIFT => "Ctrl+Alt+Shift",
        _ => "None",
    };

    public static int StringToModifier(string s) => s switch
    {
        "Alt" => MOD_ALT,
        "Ctrl" => MOD_CONTROL,
        "Shift" => MOD_SHIFT,
        "Ctrl+Alt" => MOD_CONTROL | MOD_ALT,
        "Ctrl+Shift" => MOD_CONTROL | MOD_SHIFT,
        "Alt+Shift" => MOD_ALT | MOD_SHIFT,
        "Ctrl+Alt+Shift" => MOD_CONTROL | MOD_ALT | MOD_SHIFT,
        _ => 0,
    };

    public static string KeyToString(int vk) => vk switch
    {
        >= 0x70 and <= 0x7B => $"F{vk - 0x6F}",
        >= 0x41 and <= 0x5A => ((char)vk).ToString(),
        _ => $"0x{vk:X2}",
    };

    public static int StringToKey(string s)
    {
        if (s.StartsWith("F") && int.TryParse(s[1..], out int fn) && fn is >= 1 and <= 12)
            return 0x6F + fn;
        if (s.Length == 1 && char.IsAsciiLetterUpper(s[0]))
            return s[0];
        return 0;
    }

    public string ToggleAppDisplay => $"{ModifierToString(ToggleAppModifier)} + {KeyToString(ToggleAppKey)}";
    public string ToggleTargetDisplay => $"{ModifierToString(ToggleTargetModifier)} + {KeyToString(ToggleTargetKey)}";
}
