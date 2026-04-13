// Created by Phạm Duy – Giải pháp tự động hóa thông minh.

namespace SmartMacroAI.Core;

/// <summary>Gaussian sampling (Box–Muller) and delay helpers for human-like timing.</summary>
public static class MathUtils
{
    /// <summary>Standard normal N(0,1) via Box–Muller.</summary>
    public static double NextGaussian(Random rng)
    {
        double u1 = 1.0 - rng.NextDouble();
        double u2 = 1.0 - rng.NextDouble();
        return Math.Sqrt(-2.0 * Math.Log(u1)) * Math.Cos(2 * Math.PI * u2);
    }

    /// <summary>
    /// Delay milliseconds: <paramref name="baseMs"/> + Gaussian(0, <paramref name="varianceMs"/>),
    /// clamped to [<paramref name="baseMs"/> * 0.5, <paramref name="baseMs"/> * 2.0] when <paramref name="baseMs"/> &gt; 0.
    /// </summary>
    public static int SampleHumanDelayMs(Random rng, int baseMs, int varianceMs)
    {
        if (baseMs <= 0)
            return 0;

        double v = Math.Max(0, varianceMs);
        double raw = baseMs + NextGaussian(rng) * v;
        double lo = baseMs * 0.5;
        double hi = baseMs * 2.0;
        int clamped = (int)Math.Round(Math.Clamp(raw, lo, hi));
        return Math.Max(0, clamped);
    }
}
