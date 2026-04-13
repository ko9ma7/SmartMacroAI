// Created by Phạm Duy – Giải pháp tự động hóa thông minh.

namespace SmartMacroAI.Core;

/// <summary>Thrown when Windows OCR does not complete within the configured timeout.</summary>
public sealed class OcrTimeoutException : Exception
{
    public OcrTimeoutException(string message)
        : base(message)
    {
    }
}
