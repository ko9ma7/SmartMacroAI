// Created by Phạm Duy – Giải pháp tự động hóa thông minh.

namespace SmartMacroAI.Models;

/// <summary>
/// Defines how and when a macro should run on a schedule.
/// Supports: one-time, daily, interval, and on-startup modes.
/// </summary>
public class ScheduleSettings
{
    /// <summary>When true, the scheduler is active for this macro.</summary>
    public bool Enabled { get; set; } = false;

    /// <summary>
    /// Scheduling mode. One of: "Once" | "Daily" | "Interval" | "OnStartup"
    /// </summary>
    public string Mode { get; set; } = "Once";

    /// <summary>
    /// Time of day to run (used by "Daily" mode). Default = 08:00.
    /// </summary>
    public TimeSpan RunAt { get; set; } = TimeSpan.FromHours(8);

    /// <summary>
    /// Repeat interval in minutes (used by "Interval" mode).
    /// </summary>
    public int IntervalMinutes { get; set; } = 30;

    /// <summary>
    /// One-time scheduled run (used by "Once" mode).
    /// </summary>
    public DateTime? RunOnce { get; set; }

    /// <summary>
    /// When true, the macro also runs when the application starts.
    /// </summary>
    public bool RunOnStartup { get; set; } = false;
}
