using System.Text.Json.Serialization;

namespace SmartMacroAI.Models;

/// <summary>
/// Base class for every action that can appear in a macro workflow.
/// Uses .NET 8 polymorphic JSON serialization so save/load preserves the concrete type.
/// </summary>
[JsonPolymorphic(TypeDiscriminatorPropertyName = "$type")]
[JsonDerivedType(typeof(ClickAction), "Click")]
[JsonDerivedType(typeof(WaitAction), "Wait")]
[JsonDerivedType(typeof(TypeAction), "Type")]
[JsonDerivedType(typeof(IfImageAction), "IfImage")]
[JsonDerivedType(typeof(IfTextAction), "IfText")]
[JsonDerivedType(typeof(WebNavigateAction), "WebNavigate")]
[JsonDerivedType(typeof(WebClickAction), "WebClick")]
[JsonDerivedType(typeof(WebTypeAction), "WebType")]
public abstract class MacroAction
{
    public string DisplayName { get; set; } = string.Empty;
    public bool IsEnabled { get; set; } = true;

    /// <summary>
    /// Per-action notes the user can attach in the editor.
    /// </summary>
    public string? Comment { get; set; }
}

/// <summary>
/// Sends a non-invasive left-click at (X, Y) client-coordinates on the target HWND.
/// Uses PostMessage — the physical cursor is never moved.
/// </summary>
public class ClickAction : MacroAction
{
    public int X { get; set; }
    public int Y { get; set; }

    /// <summary>
    /// When true, sends WM_RBUTTONDOWN/UP instead of WM_LBUTTONDOWN/UP.
    /// </summary>
    public bool IsRightClick { get; set; }

    public ClickAction()
    {
        DisplayName = "Click";
    }
}

/// <summary>
/// Pauses macro execution for the specified duration.
/// Implemented with <c>await Task.Delay</c> — never blocks the UI thread.
/// </summary>
public class WaitAction : MacroAction
{
    public int Milliseconds { get; set; } = 1000;

    public WaitAction()
    {
        DisplayName = "Wait";
    }
}

/// <summary>
/// Types text into the target HWND using WM_CHAR messages.
/// Non-invasive — the physical keyboard is not touched.
/// </summary>
public class TypeAction : MacroAction
{
    public string Text { get; set; } = string.Empty;

    /// <summary>
    /// Optional inter-keystroke delay (ms) to simulate human typing speed.
    /// </summary>
    public int KeyDelayMs { get; set; }

    public TypeAction()
    {
        DisplayName = "Type Text";
    }
}

/// <summary>
/// Conditional: captures the target window in the background (PrintWindow),
/// runs OpenCV template matching, and executes <see cref="ThenActions"/> only if
/// the template image is found above <see cref="Threshold"/>.
/// </summary>
public class IfImageAction : MacroAction
{
    public string ImagePath { get; set; } = string.Empty;

    /// <summary>
    /// Match confidence threshold (0.0 – 1.0). Default 0.8 = 80 %.
    /// </summary>
    public double Threshold { get; set; } = 0.8;

    /// <summary>
    /// Actions executed when the image IS found.
    /// </summary>
    public List<MacroAction> ThenActions { get; set; } = [];

    /// <summary>
    /// Actions executed when the image is NOT found (optional).
    /// </summary>
    public List<MacroAction> ElseActions { get; set; } = [];

    /// <summary>
    /// If true, the click coordinates are auto-set to the center of the matched region.
    /// </summary>
    public bool ClickOnFound { get; set; }

    public IfImageAction()
    {
        DisplayName = "IF Image Found";
    }
}

/// <summary>
/// Conditional: captures the target window, runs OCR (Tesseract),
/// and executes <see cref="ThenActions"/> only if the specified text is detected.
/// </summary>
public class IfTextAction : MacroAction
{
    public string Text { get; set; } = string.Empty;

    /// <summary>
    /// When true, the OCR comparison is case-insensitive.
    /// </summary>
    public bool IgnoreCase { get; set; } = true;

    /// <summary>
    /// When true, matches if the OCR result *contains* the text (substring match).
    /// When false, requires an exact match of the full OCR output.
    /// </summary>
    public bool PartialMatch { get; set; } = true;

    public List<MacroAction> ThenActions { get; set; } = [];
    public List<MacroAction> ElseActions { get; set; } = [];

    public IfTextAction()
    {
        DisplayName = "IF Text Found";
    }
}

/// <summary>
/// Opens a URL in the Playwright-controlled browser (visible window).
/// Hybrid workflows: use alongside Win32 desktop actions in the same script.
/// </summary>
public class WebNavigateAction : MacroAction
{
    public string Url { get; set; } = string.Empty;

    public WebNavigateAction()
    {
        DisplayName = "Web: Navigate";
    }
}

/// <summary>
/// Clicks an element via Playwright. <see cref="CssSelector"/> accepts CSS or XPath (xpath=...).
/// </summary>
public class WebClickAction : MacroAction
{
    public string CssSelector { get; set; } = string.Empty;

    public WebClickAction()
    {
        DisplayName = "Web: Click";
    }
}

/// <summary>
/// Fills an input via Playwright (clears then types).
/// </summary>
public class WebTypeAction : MacroAction
{
    public string CssSelector { get; set; } = string.Empty;
    public string TextToType { get; set; } = string.Empty;

    public WebTypeAction()
    {
        DisplayName = "Web: Type";
    }
}
