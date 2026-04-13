// Created by Phạm Duy – Giải pháp tự động hóa thông minh.

namespace SmartMacroAI.Core;

/// <summary>Mutable counters for <see cref="BehaviorRandomizer"/> (one per macro run).</summary>
public sealed class BehaviorRandomizerState
{
    public int ActionOrdinal { get; set; }

    public int StepsSinceLastMicroPause { get; set; }

    public DateTimeOffset SessionStartedUtc { get; set; } = DateTimeOffset.UtcNow;

    public bool SessionLongBreakTaken { get; set; }

    public void Reset()
    {
        ActionOrdinal = 0;
        StepsSinceLastMicroPause = 0;
        SessionStartedUtc = DateTimeOffset.UtcNow;
        SessionLongBreakTaken = false;
    }
}
