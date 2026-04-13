// Created by Phạm Duy – Giải pháp tự động hóa thông minh.

using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Threading;

namespace SmartMacroAI.Core;

public enum ModuleCategory
{
    KnownFramework,
    AppLocal,
    Unknown,
}

/// <summary>
/// Decoy window title rotation, optional capture exclusion, and a 3-tier loaded-module audit.
/// Fail-safe: never throws to callers.
/// </summary>
public sealed class ModuleAuditService
{
    private static readonly Lazy<ModuleAuditService> Lazy = new(() => new ModuleAuditService());

    public static ModuleAuditService Instance => Lazy.Value;

    /// <summary>Optional sink for audit lines; when null, <see cref="Console.WriteLine"/> is used for unknown modules.</summary>
    public static event Action<string>? AuditLog;

    private DispatcherTimer? _titleTimer;

    private Window? _window;

    public void AttachWindow(Window window)
    {
        try
        {
            _window = window;
        }
        catch
        {
            /* ignore */
        }
    }

    public void StartTitleRandomizerIfEnabled(SmartMacroAI.Localization.AppSettings app)
    {
        try
        {
            _titleTimer?.Stop();
            if (!app.AntiDetectionEnabled || app.AntiDetectionDecoyTitles.Count == 0 || _window is null)
                return;

            _titleTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(Random.Shared.Next(30, 61)) };
            _titleTimer.Tick += (_, _) =>
            {
                try
                {
                    if (_window is null)
                        return;
                    string pick = app.AntiDetectionDecoyTitles[Random.Shared.Next(app.AntiDetectionDecoyTitles.Count)];
                    _window.Title = string.IsNullOrWhiteSpace(pick) ? "SmartMacroAI" : pick.Trim();
                    _titleTimer!.Interval = TimeSpan.FromSeconds(Random.Shared.Next(30, 61));
                }
                catch
                {
                    /* ignore */
                }
            };
            _titleTimer.Start();
        }
        catch
        {
            /* ignore */
        }
    }

    public void StopTitleRandomizer()
    {
        try
        {
            _titleTimer?.Stop();
            _titleTimer = null;
        }
        catch
        {
            /* ignore */
        }
    }

    public static void ApplyExcludeFromCapture(IntPtr hwnd, bool enable)
    {
        if (hwnd == IntPtr.Zero)
            return;
        try
        {
            _ = NativeMethods.SetWindowDisplayAffinity(
                hwnd,
                enable ? NativeMethods.WDA_EXCLUDEFROMCAPTURE : NativeMethods.WDA_NONE);
        }
        catch
        {
            /* ignore */
        }
    }

    /// <summary>Tier 1–3 plus app directory: true when module is not treated as unknown/suspicious.</summary>
    public static bool IsModuleSafe(ProcessModule module) =>
        ClassifyModule(module) != ModuleCategory.Unknown;

    /// <summary>
    /// Tier 1: under Windows directory. Tier 2: under Program Files\dotnet. Tier 3: known framework file name prefix.
    /// App binaries under <see cref="AppDomain.CurrentDomain.BaseDirectory"/> are <see cref="ModuleCategory.AppLocal"/>.
    /// </summary>
    public static ModuleCategory ClassifyModule(ProcessModule module)
    {
        string? path = module.FileName;
        if (string.IsNullOrWhiteSpace(path))
            return ModuleCategory.Unknown;

        string full;
        try
        {
            full = Path.GetFullPath(path);
        }
        catch
        {
            return ModuleCategory.Unknown;
        }

        string normBase;
        try
        {
            normBase = AppendDirSep(Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory));
        }
        catch
        {
            normBase = AppendDirSep(AppDomain.CurrentDomain.BaseDirectory);
        }

        if (full.StartsWith(normBase, StringComparison.OrdinalIgnoreCase)
            || string.Equals(full, normBase.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar), StringComparison.OrdinalIgnoreCase))
            return ModuleCategory.AppLocal;

        string winDir = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
        if (!string.IsNullOrEmpty(winDir))
        {
            string winRoot = AppendDirSep(Path.GetFullPath(winDir));
            if (full.StartsWith(winRoot, StringComparison.OrdinalIgnoreCase))
                return ModuleCategory.KnownFramework;
        }

        string dotnetDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
            "dotnet");
        if (!string.IsNullOrEmpty(dotnetDir))
        {
            try
            {
                string dotnetRoot = AppendDirSep(Path.GetFullPath(dotnetDir));
                if (full.StartsWith(dotnetRoot, StringComparison.OrdinalIgnoreCase))
                    return ModuleCategory.KnownFramework;
            }
            catch
            {
                /* ignore */
            }
        }

        string fileName = Path.GetFileName(full);
        if (fileName.Length == 0)
            return ModuleCategory.Unknown;

        ReadOnlySpan<string> prefixes =
        [
            "System.",
            "Presentation",
            "Microsoft.",
            "WindowsBase",
            "hostfxr",
            "DirectWrite",
            "clr",
            "mscor",
        ];

        foreach (string p in prefixes)
        {
            if (fileName.StartsWith(p, StringComparison.OrdinalIgnoreCase))
                return ModuleCategory.KnownFramework;
        }

        return ModuleCategory.Unknown;
    }

    private static string AppendDirSep(string dir)
    {
        if (string.IsNullOrEmpty(dir))
            return dir;
        char c = dir[^1];
        if (c == Path.DirectorySeparatorChar || c == Path.AltDirectorySeparatorChar)
            return dir;
        return dir + Path.DirectorySeparatorChar;
    }

    private static void EmitAudit(string message)
    {
        Action<string>? h = AuditLog;
        if (h is null)
            Console.WriteLine(message);
        else
            h(message);
    }

    private static string DescribeFailedTiers(string fullPath, string fileName, string winDir, string dotnetDirCombined)
    {
        var parts = new List<string>(4);
        string winRoot = string.IsNullOrEmpty(winDir) ? "" : AppendDirSep(Path.GetFullPath(winDir));
        if (string.IsNullOrEmpty(winRoot) || !fullPath.StartsWith(winRoot, StringComparison.OrdinalIgnoreCase))
            parts.Add("Tier1 (not under Windows folder)");

        string dotnetRoot = "";
        try
        {
            if (!string.IsNullOrEmpty(dotnetDirCombined))
                dotnetRoot = AppendDirSep(Path.GetFullPath(dotnetDirCombined));
        }
        catch
        {
            dotnetRoot = "";
        }

        if (string.IsNullOrEmpty(dotnetRoot) || !fullPath.StartsWith(dotnetRoot, StringComparison.OrdinalIgnoreCase))
            parts.Add("Tier2 (not under ProgramFiles\\dotnet)");

        parts.Add("Tier3 (filename has no known framework prefix: System.*, Presentation*, Microsoft.*, WindowsBase*, hostfxr*, DirectWrite*, clr*, mscor*)");

        return string.Join("; ", parts);
    }

    /// <summary>Scans loaded modules; alerts once for modules classified as <see cref="ModuleCategory.Unknown"/>.</summary>
    public static void ScanForeignModulesOnStartupIfEnabled(Action<string> alertSink)
    {
        try
        {
            var app = SmartMacroAI.Localization.AppSettings.Load();
            if (!app.AntiDetectionEnabled || !app.AntiDetectionHookScanOnStartup)
                return;

            using Process p = Process.GetCurrentProcess();
            string winDir = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
            string dotnetDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                "dotnet");

            var suspicious = new List<string>();

            foreach (ProcessModule? m in p.Modules)
            {
                if (m is null)
                    continue;

                ModuleCategory cat = ClassifyModule(m);
                if (cat != ModuleCategory.Unknown)
                    continue;

                string path = m.FileName ?? "";
                string fileName = Path.GetFileName(path);
                string full = path;
                try
                {
                    full = Path.GetFullPath(path);
                }
                catch
                {
                    /* keep raw */
                }

                string failed = DescribeFailedTiers(full, fileName, winDir, dotnetDir);
                EmitAudit($"[ModuleAudit] UNKNOWN | File={fileName} | Path={full} | Failed: {failed}");

                suspicious.Add(m.ModuleName ?? fileName);
            }

            if (suspicious.Count > 0)
            {
                alertSink(
                    "Phát hiện module lạ đang theo dõi! " +
                    string.Join(", ", suspicious.Distinct(StringComparer.OrdinalIgnoreCase).Take(8)));
            }
        }
        catch (Exception ex)
        {
            alertSink?.Invoke($"[Anti-Detection] Hook scan failed: {ex.Message}");
        }
    }
}
