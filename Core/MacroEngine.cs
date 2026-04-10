using System.Drawing;
using System.IO;
using SmartMacroAI.Models;

namespace SmartMacroAI.Core;

/// <summary>
/// Asynchronous macro execution engine.
/// Runs entirely on background threads — the WPF UI thread is never blocked.
/// All window interactions go through <see cref="Win32Api"/> (PostMessage / SendMessage),
/// which means the physical mouse and keyboard are NEVER hijacked.
/// Web steps use <see cref="PlaywrightEngine"/> (headful browser) in parallel with desktop actions.
/// </summary>
public sealed class MacroEngine
{
    private PlaywrightEngine? _playwrightEngine;
    // ═══════════════════════════════════════════════
    //  EVENTS — for UI progress/logging
    // ═══════════════════════════════════════════════

    public event Action<string>? Log;
    public event Action<MacroAction, int>? ActionStarted;
    public event Action<MacroAction, int>? ActionCompleted;
    public event Action<int, int>? IterationStarted;
    public event Action? ExecutionFinished;
    public event Action<Exception>? ExecutionFaulted;

    private void OnLog(string message) => Log?.Invoke(message);

    // ═══════════════════════════════════════════════
    //  PUBLIC API
    // ═══════════════════════════════════════════════

    /// <summary>
    /// Executes a <see cref="MacroScript"/> against the window whose title
    /// matches <see cref="MacroScript.TargetWindowTitle"/>.
    /// Runs asynchronously on background threads — safe to call from the UI.
    /// </summary>
    public async Task ExecuteScriptAsync(MacroScript script, CancellationToken token)
    {
        ArgumentNullException.ThrowIfNull(script);

        if (string.IsNullOrWhiteSpace(script.TargetWindowTitle))
            throw new ArgumentException(
                "TargetWindowTitle must be set before execution.", nameof(script));

        IntPtr hwnd = Win32Api.FindWindowByPartialTitle(script.TargetWindowTitle);
        if (hwnd == IntPtr.Zero)
            throw new InvalidOperationException(
                $"Target window not found: \"{script.TargetWindowTitle}\". " +
                "If the window is hidden via Stealth, use the pre-resolved HWND overload.");

        string windowTitle = Win32Api.GetWindowTitle(hwnd);
        OnLog($"Target acquired: \"{windowTitle}\" (HWND=0x{hwnd:X})");

        try
        {
            await RunLoopAsync(script, hwnd, token);
            OnLog("Execution completed successfully.");
            ExecutionFinished?.Invoke();
        }
        catch (OperationCanceledException)
        {
            OnLog("Execution cancelled by user.");
            throw;
        }
        catch (Exception ex)
        {
            OnLog($"Execution faulted: {ex.Message}");
            ExecutionFaulted?.Invoke(ex);
            throw;
        }
    }

    /// <summary>
    /// Overload that accepts a pre-resolved HWND (useful when the
    /// UI has already identified the target window).
    /// </summary>
    public async Task ExecuteScriptAsync(MacroScript script, IntPtr hwnd, CancellationToken token)
    {
        ArgumentNullException.ThrowIfNull(script);

        if (hwnd == IntPtr.Zero || !Win32Api.IsWindow(hwnd))
            throw new ArgumentException("Invalid window handle.", nameof(hwnd));

        string windowTitle = Win32Api.GetWindowTitle(hwnd);
        OnLog($"Target acquired: \"{windowTitle}\" (HWND=0x{hwnd:X})");

        try
        {
            await RunLoopAsync(script, hwnd, token);
            OnLog("Execution completed successfully.");
            ExecutionFinished?.Invoke();
        }
        catch (OperationCanceledException)
        {
            OnLog("Execution cancelled by user.");
            throw;
        }
        catch (Exception ex)
        {
            OnLog($"Execution faulted: {ex.Message}");
            ExecutionFaulted?.Invoke(ex);
            throw;
        }
    }

    // ═══════════════════════════════════════════════
    //  REPEAT LOOP
    // ═══════════════════════════════════════════════

    private async Task RunLoopAsync(MacroScript script, IntPtr hwnd, CancellationToken token)
    {
        try
        {
            bool infinite = script.RepeatCount <= 0;
            int totalIterations = infinite ? int.MaxValue : script.RepeatCount;

            for (int i = 1; i <= totalIterations; i++)
            {
                token.ThrowIfCancellationRequested();

                if (!Win32Api.IsWindow(hwnd))
                    throw new InvalidOperationException(
                        "Target window was closed during execution.");

                string label = infinite ? $"#{i} (infinite)" : $"#{i}/{script.RepeatCount}";
                OnLog($"── Iteration {label} ──");
                IterationStarted?.Invoke(i, script.RepeatCount);

                await ExecuteActionsAsync(script.Actions, hwnd, token);

                bool hasMore = i < totalIterations;
                if (hasMore && script.IntervalMinutes > 0)
                {
                    OnLog($"── Waiting {script.IntervalMinutes} min before next iteration ──");
                    await Task.Delay(TimeSpan.FromMinutes(script.IntervalMinutes), token);
                }
            }
        }
        finally
        {
            if (_playwrightEngine is not null)
            {
                try
                {
                    await _playwrightEngine.DisposeAsync().ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    OnLog($"Web engine dispose: {ex.Message}");
                }

                _playwrightEngine = null;
            }
        }
    }

    // ═══════════════════════════════════════════════
    //  ACTION DISPATCHER  (recursive for nested IF blocks)
    // ═══════════════════════════════════════════════

    private async Task ExecuteActionsAsync(
        List<MacroAction> actions, IntPtr hwnd, CancellationToken token)
    {
        for (int idx = 0; idx < actions.Count; idx++)
        {
            token.ThrowIfCancellationRequested();

            var action = actions[idx];

            if (!action.IsEnabled)
            {
                OnLog($"  [{idx}] {action.DisplayName} — SKIPPED (disabled)");
                continue;
            }

            ActionStarted?.Invoke(action, idx);

            switch (action)
            {
                case ClickAction click:
                    await ExecuteClickAsync(click, hwnd);
                    break;

                case WaitAction wait:
                    await ExecuteWaitAsync(wait, token);
                    break;

                case TypeAction type:
                    await ExecuteTypeAsync(type, hwnd, token);
                    break;

                case IfImageAction ifImage:
                    await ExecuteIfImageAsync(ifImage, hwnd, token);
                    break;

                case IfTextAction ifText:
                    await ExecuteIfTextAsync(ifText, hwnd, token);
                    break;

                case WebNavigateAction webNav:
                    await ExecuteWebNavigateAsync(webNav, token);
                    break;

                case WebClickAction webClick:
                    await ExecuteWebClickAsync(webClick, token);
                    break;

                case WebTypeAction webType:
                    await ExecuteWebTypeAsync(webType, token);
                    break;

                default:
                    OnLog($"  [{idx}] Unknown action type: {action.GetType().Name}");
                    break;
            }

            ActionCompleted?.Invoke(action, idx);
        }
    }

    // ═══════════════════════════════════════════════
    //  ACTION HANDLERS
    // ═══════════════════════════════════════════════

    private async Task ExecuteClickAsync(ClickAction click, IntPtr hwnd)
    {
        if (click.IsRightClick)
        {
            OnLog($"  ControlRightClick({click.X}, {click.Y})");
            await Win32Api.ControlRightClickAsync(hwnd, click.X, click.Y);
        }
        else
        {
            OnLog($"  ControlClick({click.X}, {click.Y})");
            await Win32Api.ControlClickAsync(hwnd, click.X, click.Y);
        }
    }

    private async Task ExecuteWaitAsync(WaitAction wait, CancellationToken token)
    {
        OnLog($"  Wait {wait.Milliseconds}ms");
        await Task.Delay(wait.Milliseconds, token);
    }

    private async Task ExecuteTypeAsync(TypeAction type, IntPtr hwnd, CancellationToken token)
    {
        OnLog($"  TypeText \"{Truncate(type.Text, 40)}\" (delay={type.KeyDelayMs}ms)");

        if (type.KeyDelayMs <= 0)
        {
            Win32Api.ControlSendText(hwnd, type.Text);
        }
        else
        {
            foreach (char c in type.Text)
            {
                token.ThrowIfCancellationRequested();
                Win32Api.PostMessage(hwnd, Win32Api.WM_CHAR, (IntPtr)c, IntPtr.Zero);
                await Task.Delay(type.KeyDelayMs, token);
            }
        }
    }

    private async Task ExecuteIfImageAsync(
        IfImageAction ifImage, IntPtr hwnd, CancellationToken token)
    {
        OnLog($"  IfImageFound \"{Path.GetFileName(ifImage.ImagePath)}\" " +
              $"(threshold={ifImage.Threshold:P0})");

        Point? match = null;
        try
        {
            match = VisionEngine.FindImageOnWindow(hwnd, ifImage.ImagePath, ifImage.Threshold);
        }
        catch (Exception ex)
        {
            OnLog($"    ⚠ Vision error: {ex.Message}");
        }

        if (match.HasValue)
        {
            OnLog($"    → FOUND at ({match.Value.X}, {match.Value.Y})");

            if (ifImage.ClickOnFound)
            {
                OnLog($"    → ControlClick({match.Value.X}, {match.Value.Y})");
                await Win32Api.ControlClickAsync(hwnd, match.Value.X, match.Value.Y);
            }

            if (ifImage.ThenActions.Count > 0)
            {
                OnLog($"    → Executing {ifImage.ThenActions.Count} THEN action(s)");
                await ExecuteActionsAsync(ifImage.ThenActions, hwnd, token);
            }
        }
        else
        {
            OnLog("    → NOT FOUND");

            if (ifImage.ElseActions.Count > 0)
            {
                OnLog($"    → Executing {ifImage.ElseActions.Count} ELSE action(s)");
                await ExecuteActionsAsync(ifImage.ElseActions, hwnd, token);
            }
        }
    }

    private async Task ExecuteWebNavigateAsync(WebNavigateAction nav, CancellationToken token)
    {
        if (string.IsNullOrWhiteSpace(nav.Url))
        {
            OnLog("  WebNavigate — SKIPPED (empty URL)");
            return;
        }

        _playwrightEngine ??= new PlaywrightEngine();
        OnLog($"  WebNavigate: {nav.Url}");
        await _playwrightEngine.MapsAsync(nav.Url.Trim(), token).ConfigureAwait(false);
    }

    private async Task ExecuteWebClickAsync(WebClickAction click, CancellationToken token)
    {
        if (string.IsNullOrWhiteSpace(click.CssSelector))
        {
            OnLog("  WebClick — SKIPPED (empty selector)");
            return;
        }

        _playwrightEngine ??= new PlaywrightEngine();
        await _playwrightEngine.EnsureBrowserStartedAsync(token).ConfigureAwait(false);
        OnLog($"  WebClick: {click.CssSelector}");
        await _playwrightEngine.ClickSelectorAsync(click.CssSelector.Trim(), token).ConfigureAwait(false);
    }

    private async Task ExecuteWebTypeAsync(WebTypeAction type, CancellationToken token)
    {
        if (string.IsNullOrWhiteSpace(type.CssSelector))
        {
            OnLog("  WebType — SKIPPED (empty selector)");
            return;
        }

        _playwrightEngine ??= new PlaywrightEngine();
        await _playwrightEngine.EnsureBrowserStartedAsync(token).ConfigureAwait(false);
        OnLog($"  WebType: {type.CssSelector} ← \"{Truncate(type.TextToType, 40)}\"");
        await _playwrightEngine.TypeSelectorAsync(type.CssSelector.Trim(), type.TextToType ?? "", token)
            .ConfigureAwait(false);
    }

    private async Task ExecuteIfTextAsync(
        IfTextAction ifText, IntPtr hwnd, CancellationToken token)
    {
        OnLog($"  IfTextFound \"{Truncate(ifText.Text, 30)}\" " +
              $"(ignoreCase={ifText.IgnoreCase}, partial={ifText.PartialMatch})");

        string ocrResult;
        try
        {
            ocrResult = VisionEngine.ExtractTextFromWindow(hwnd);
        }
        catch (Exception ex)
        {
            OnLog($"    ⚠ OCR error: {ex.Message}");
            ocrResult = string.Empty;
        }

        bool found;
        if (ifText.PartialMatch)
        {
            found = ocrResult.Contains(
                ifText.Text,
                ifText.IgnoreCase
                    ? StringComparison.OrdinalIgnoreCase
                    : StringComparison.Ordinal);
        }
        else
        {
            found = string.Equals(
                ocrResult.Trim(), ifText.Text.Trim(),
                ifText.IgnoreCase
                    ? StringComparison.OrdinalIgnoreCase
                    : StringComparison.Ordinal);
        }

        if (found)
        {
            OnLog("    → TEXT FOUND");

            if (ifText.ThenActions.Count > 0)
            {
                OnLog($"    → Executing {ifText.ThenActions.Count} THEN action(s)");
                await ExecuteActionsAsync(ifText.ThenActions, hwnd, token);
            }
        }
        else
        {
            OnLog("    → TEXT NOT FOUND");

            if (ifText.ElseActions.Count > 0)
            {
                OnLog($"    → Executing {ifText.ElseActions.Count} ELSE action(s)");
                await ExecuteActionsAsync(ifText.ElseActions, hwnd, token);
            }
        }
    }

    // ═══════════════════════════════════════════════
    //  UTIL
    // ═══════════════════════════════════════════════

    private static string Truncate(string value, int maxLength) =>
        value.Length <= maxLength ? value : string.Concat(value.AsSpan(0, maxLength), "…");
}
