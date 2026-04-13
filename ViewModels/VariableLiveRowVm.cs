// Created by Phạm Duy – Giải pháp tự động hóa thông minh.

namespace SmartMacroAI.ViewModels;

/// <summary>One row in the Dashboard live variables grid.</summary>
public sealed class VariableLiveRowVm
{
    public string Name { get; set; } = "";
    public string Value { get; set; } = "";
    public string Source { get; set; } = "";
}
