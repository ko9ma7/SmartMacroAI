namespace SmartMacroAI.Models;

/// <summary>
/// Record of a single macro run (for run history).
/// Created by Phạm Duy – Giải pháp tự động hóa thông minh.
/// </summary>
public sealed class MacroRunRecord
{
    public string     MacroName       { get; set; } = "";
    public string     MacroFile       { get; set; } = "";
    public DateTime   StartTime       { get; set; }
    public DateTime?  EndTime         { get; set; }
    public bool       Success         { get; set; }
    public int        TotalSteps      { get; set; }
    public int        CompletedSteps  { get; set; }
    public string     ErrorMessage    { get; set; } = "";
    public string?    ScreenshotPath  { get; set; }
    public string     LogSnapshot     { get; set; } = "";

    public string Duration =>
        EndTime.HasValue
            ? (EndTime.Value - StartTime).ToString(@"mm\:ss")
            : "...";

    public string StatusIcon => Success ? "✅" : "❌";

    public string StatusColor => Success ? "#A6E3A1" : "#F38BA8";

    public string StepsSummary =>
        TotalSteps > 0
            ? $"{CompletedSteps}/{TotalSteps}"
            : "-";

    public bool HasScreenshot => !string.IsNullOrWhiteSpace(ScreenshotPath);
}
