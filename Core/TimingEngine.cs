// Created by Phạm Duy – Giải pháp tự động hóa thông minh.

using SmartMacroAI.Localization;

namespace SmartMacroAI.Core;

/// <summary>
/// Async delays with Gaussian jitter, optional fatigue (+5% every 50 steps, max +40%),
/// and light CPU touch to avoid flat 0% idle spikes when enabled in settings.
/// </summary>
public sealed class TimingEngine : ITimingEngine
{
    private readonly Random _rng = new();

    private int _macroStepCount;

    private double _fatigueMultiplier = 1.0;

    /// <summary>Exposed for unit tests (1.0–1.40 when fatigue is active).</summary>
    public double CurrentFatigueMultiplier => _fatigueMultiplier;

    public void ResetSession()
    {
        _macroStepCount = 0;
        _fatigueMultiplier = 1.0;
    }

    public void NotifyMacroStepCompleted()
    {
        _macroStepCount++;
        var app = AppSettings.Load();
        if (!app.AntiDetectionEnabled || !app.AntiDetectionFatigueEnabled)
            return;

        if (_macroStepCount > 0 && _macroStepCount % 50 == 0 && _fatigueMultiplier < 1.40 - 1e-6)
        {
            _fatigueMultiplier = Math.Min(1.40, _fatigueMultiplier + 0.05);
        }
    }

    public async Task WaitAsync(int baseMilliseconds, int varianceMilliseconds, CancellationToken cancellationToken = default)
    {
        if (baseMilliseconds <= 0)
        {
            await Task.Yield();
            return;
        }

        var app = AppSettings.Load();
        if (!app.AntiDetectionEnabled)
        {
            await Task.Delay(baseMilliseconds, cancellationToken).ConfigureAwait(false);
            return;
        }

        try
        {
            // Long session breaks: avoid heavy Gaussian clamp distortion; small uniform jitter only.
            if (baseMilliseconds >= 120_000)
            {
                int jitter = _rng.Next(-Math.Max(1, baseMilliseconds / 60), Math.Max(1, baseMilliseconds / 60) + 1);
                int msLong = Math.Max(1000, baseMilliseconds + jitter);
                await Task.Delay(msLong, cancellationToken).ConfigureAwait(false);
                if (app.AntiDetectionCpuIdleTweak)
                    RunLightCpuTouch(_rng.Next(800, 2000));
                return;
            }

            int ms = MathUtils.SampleHumanDelayMs(_rng, baseMilliseconds, Math.Max(0, varianceMilliseconds));
            double fatigue = app.AntiDetectionFatigueEnabled ? _fatigueMultiplier : 1.0;
            ms = (int)Math.Round(ms * fatigue);
            ms = Math.Max(0, ms);
            await Task.Delay(ms, cancellationToken).ConfigureAwait(false);

            if (app.AntiDetectionCpuIdleTweak && ms > 0)
                RunLightCpuTouch(_rng.Next(400, 1200));
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch
        {
            await Task.Delay(baseMilliseconds, cancellationToken).ConfigureAwait(false);
        }
    }

    /// <summary>Short busy work (no await) to avoid a perfectly flat scheduler profile.</summary>
    private static void RunLightCpuTouch(int iterations)
    {
        try
        {
            string sink = "";
            for (int i = 0; i < iterations; i++)
                sink = string.Concat(sink, (char)('0' + (i % 10)));
            GC.KeepAlive(sink);
        }
        catch
        {
            /* fail-safe */
        }
    }
}
