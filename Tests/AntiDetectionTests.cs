// Created by Phạm Duy – Giải pháp tự động hóa thông minh.

using SmartMacroAI.Core;
using Xunit;

namespace SmartMacroAI.Tests;

public sealed class AntiDetectionTests
{
    [Fact]
    public void MathUtils_SampledDelays_AreClampedAroundBase()
    {
        var rng = new Random(42);
        const int baseMs = 200;
        const int variance = 40;
        var samples = new int[800];
        for (int i = 0; i < samples.Length; i++)
            samples[i] = MathUtils.SampleHumanDelayMs(rng, baseMs, variance);

        Assert.All(samples, v =>
        {
            Assert.InRange(v, (int)(baseMs * 0.5) - 1, (int)(baseMs * 2.0) + 1);
        });

        double mean = samples.Average();
        Assert.InRange(mean, 150, 250);
    }

    [Fact]
    public void BehaviorRandomizer_MisclickRate_ApproximatesConfiguredPercent()
    {
        var rng = new Random(17);
        const int misPct = 5;
        int hits = 0;
        const int n = 4000;
        for (int i = 0; i < n; i++)
        {
            if (BehaviorRandomizer.RollMisclick(rng, misPct))
                hits++;
        }

        double p = hits / (double)n;
        Assert.InRange(p, 0.03, 0.08);
    }

    [Fact]
    public async Task MacroRunner_WaitAsync_CompletesUnderCancellation()
    {
        var runner = new MacroRunner(new TimingEngine());
        using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(3));
        await runner.Timing.WaitAsync(8, 3, cts.Token);
    }

    [Fact]
    public void TimingEngine_NotifySteps_DoesNotThrow()
    {
        var eng = new TimingEngine();
        for (int i = 0; i < 160; i++)
            eng.NotifyMacroStepCompleted();
        Assert.InRange(eng.CurrentFatigueMultiplier, 1.0, 1.45);
    }
}
