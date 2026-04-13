// Created by Phạm Duy – Giải pháp tự động hóa thông minh.

namespace SmartMacroAI.Core;

/// <summary>
/// Facade for macro timing and anti-detection pacing. Injected into <see cref="MacroEngine"/>
/// (there is no separate legacy runner — this type centralizes <see cref="ITimingEngine"/> usage).
/// </summary>
public sealed class MacroRunner
{
    public MacroRunner(ITimingEngine? timing = null)
    {
        Timing = timing ?? new TimingEngine();
    }

    public ITimingEngine Timing { get; }

    /// <summary>Gaussian delay; <paramref name="varianceDivisor"/> divides <paramref name="baseMs"/> for default spread.</summary>
    public Task WaitAsync(int baseMs, CancellationToken ct, int varianceDivisor = 5)
    {
        int v = baseMs <= 0 ? 0 : Math.Max(1, baseMs / Math.Max(1, varianceDivisor));
        return Timing.WaitAsync(baseMs, v, ct);
    }
}
