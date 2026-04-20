// Created by Phạm Duy – Giải pháp tự động hóa thông minh.

using System.Collections.Concurrent;
using SmartMacroAI.Models;

namespace SmartMacroAI.Core;

/// <summary>
/// Central scheduler service that manages timed macro execution.
/// Timers are registered per-macro based on their <see cref="ScheduleSettings"/>.
/// Disposes all timers when the application closes.
/// </summary>
public static class SchedulerService
{
    private static readonly ConcurrentDictionary<string, (MacroScript Script, System.Threading.Timer Timer)> _timers = new();
    private static readonly ConcurrentDictionary<string, DateTime> _lastRunTimes = new();

    /// <summary>Event raised when a scheduled macro fires. Parameter: macro name.</summary>
    public static event Action<string>? MacroTriggered;

    /// <summary>
    /// Registers (or re-registers) a macro's schedule. Removes any existing timer first.
    /// </summary>
    public static void Register(MacroScript script, Func<MacroScript, Task> runCallback)
    {
        Unregister(script.Name);

        if (script.Schedule == null || !script.Schedule.Enabled)
            return;

        switch (script.Schedule.Mode)
        {
            case "Daily":
                RegisterDaily(script, runCallback);
                break;
            case "Interval":
                RegisterInterval(script, runCallback);
                break;
            case "Once":
                RegisterOnce(script, runCallback);
                break;
            case "OnStartup":
                // No timer needed — handled in RegisterAllOnStartup
                break;
        }
    }

    /// <summary>
    /// Runs macros that have RunOnStartup = true.
    /// </summary>
    public static void RegisterAllOnStartup(IEnumerable<MacroScript> scripts, Func<MacroScript, Task> runCallback)
    {
        foreach (var script in scripts)
        {
            if (script.Schedule?.Enabled == true && script.Schedule.RunOnStartup)
            {
                _ = runCallback(script);
                _lastRunTimes[$"OnStartup:{script.Name}"] = DateTime.Now;
            }
        }
    }

    private static void RegisterDaily(MacroScript script, Func<MacroScript, Task> callback)
    {
        var now = DateTime.Now;
        var target = DateTime.Today.Add(script.Schedule.RunAt);
        if (target <= now)
            target = target.AddDays(1);

        var delay = target - now;
        var timer = new System.Threading.Timer(
            async _ => await FireMacro(script, callback),
            null, delay, TimeSpan.FromDays(1));

        _timers[$"Daily:{script.Name}"] = (script, timer);
    }

    private static void RegisterInterval(MacroScript script, Func<MacroScript, Task> callback)
    {
        var interval = TimeSpan.FromMinutes(script.Schedule.IntervalMinutes);
        var timer = new System.Threading.Timer(
            async _ => await FireMacro(script, callback),
            null, interval, interval);

        _timers[$"Interval:{script.Name}"] = (script, timer);
    }

    private static void RegisterOnce(MacroScript script, Func<MacroScript, Task> callback)
    {
        if (!script.Schedule.RunOnce.HasValue)
            return;

        var delay = script.Schedule.RunOnce.Value - DateTime.Now;
        if (delay <= TimeSpan.Zero)
            return;

        var timer = new System.Threading.Timer(
            async _ => await FireMacro(script, callback),
            null, delay, Timeout.InfiniteTimeSpan);

        _timers[$"Once:{script.Name}"] = (script, timer);
    }

    private static async Task FireMacro(MacroScript script, Func<MacroScript, Task> callback)
    {
        var key = $"Fire:{script.Name}";
        _lastRunTimes[key] = DateTime.Now;
        MacroTriggered?.Invoke(script.Name);
        await callback(script);
    }

    /// <summary>
    /// Returns the last fire time for a given macro name, or null if never fired.
    /// </summary>
    public static DateTime? GetLastRunTime(string macroName)
    {
        foreach (var key in _lastRunTimes.Keys)
        {
            if (key.EndsWith($":{macroName}"))
                return _lastRunTimes[key];
        }
        return null;
    }

    /// <summary>
    /// Removes and disposes the timer for a named macro.
    /// </summary>
    public static void Unregister(string macroName)
    {
        foreach (var key in _timers.Keys.Where(k => k.EndsWith($":{macroName}")).ToList())
        {
            if (_timers.TryRemove(key, out var entry))
            {
                entry.Timer.Dispose();
            }
        }
    }

    /// <summary>
    /// Disposes all timers. Call this when the application closes.
    /// </summary>
    public static void UnregisterAll()
    {
        foreach (var entry in _timers.Values)
        {
            entry.Timer.Dispose();
        }
        _timers.Clear();
    }
}
