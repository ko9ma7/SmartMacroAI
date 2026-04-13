// Created by Phạm Duy – Giải pháp tự động hóa thông minh.

using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Threading;

namespace SmartMacroAI.Core;

/// <summary>
/// Decoy window title rotation, optional capture exclusion, injected-module scan, and related UI hooks.
/// Fail-safe: never throws to callers.
/// </summary>
public sealed class ProcessStealthService
{
    private static readonly Lazy<ProcessStealthService> Lazy = new(() => new ProcessStealthService());

    public static ProcessStealthService Instance => Lazy.Value;

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

    /// <summary>Scans loaded modules; alerts once for paths outside runtime + framework heuristics.</summary>
    public static void ScanForeignModulesOnStartupIfEnabled(Action<string> alertSink)
    {
        try
        {
            var app = SmartMacroAI.Localization.AppSettings.Load();
            if (!app.AntiDetectionEnabled || !app.AntiDetectionHookScanOnStartup)
                return;

            using Process p = Process.GetCurrentProcess();
            string baseDir = Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory);
            string windir = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
            var suspicious = new List<string>();

            foreach (ProcessModule? m in p.Modules)
            {
                if (m is null)
                    continue;
                string path = m.FileName ?? "";
                if (path.Length == 0)
                    continue;
                string full = Path.GetFullPath(path);
                if (full.StartsWith(baseDir, StringComparison.OrdinalIgnoreCase))
                    continue;
                if (full.StartsWith(windir, StringComparison.OrdinalIgnoreCase))
                    continue;
                if (full.Contains("Microsoft.NET", StringComparison.OrdinalIgnoreCase)
                    || full.Contains("\\Windows\\", StringComparison.OrdinalIgnoreCase))
                    continue;

                suspicious.Add(m.ModuleName ?? Path.GetFileName(path));
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
