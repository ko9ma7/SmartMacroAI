// Created by Phạm Duy – Giải pháp tự động hóa thông minh.

namespace SmartMacroAI.Core;

/// <summary>Human-like async waits (Gaussian + optional fatigue). Mockable for tests.</summary>
public interface ITimingEngine
{
    void ResetSession();

    /// <summary>Counts a completed macro action (drives fatigue tiers).</summary>
    void NotifyMacroStepCompleted();

    Task WaitAsync(int baseMilliseconds, int varianceMilliseconds, CancellationToken cancellationToken = default);
}
