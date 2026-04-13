// Created by Phạm Duy – Giải pháp tự động hóa thông minh.

using System.Drawing;
using System.IO;
using System.Drawing.Imaging;
using System.Windows;
using Windows.Globalization;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;
using Windows.Storage.Streams;

namespace SmartMacroAI.Core;

/// <summary>
/// Screen-region OCR using <see cref="OcrEngine"/> (Windows.Media.Ocr).
/// </summary>
public sealed class OcrService
{
    private readonly string _languageTag;

    /// <param name="languageTag"><c>auto</c>, <c>vi-VN</c>, or <c>en-US</c>.</param>
    public OcrService(string? languageTag = null)
    {
        _languageTag = string.IsNullOrWhiteSpace(languageTag) ? "auto" : languageTag.Trim();
    }

    /// <summary>
    /// Captures the screen rectangle and runs OCR. Clips to the virtual screen.
    /// </summary>
    /// <exception cref="OcrTimeoutException">When recognition exceeds <paramref name="timeout"/>.</exception>
    public async Task<(string Text, double AverageConfidence)> ReadTextFromRegionWithConfidenceAsync(
        Rectangle region,
        TimeSpan timeout,
        CancellationToken cancellationToken = default)
    {
        int vx = (int)SystemParameters.VirtualScreenLeft;
        int vy = (int)SystemParameters.VirtualScreenTop;
        int vw = (int)SystemParameters.VirtualScreenWidth;
        int vh = (int)SystemParameters.VirtualScreenHeight;

        var clipped = Rectangle.Intersect(
            new Rectangle(vx, vy, vw, vh),
            new Rectangle(region.X, region.Y, region.Width, region.Height));

        if (clipped.Width <= 0 || clipped.Height <= 0)
            return (string.Empty, 0);

        using var bmp = new Bitmap(clipped.Width, clipped.Height, PixelFormat.Format32bppArgb);
        using (var g = Graphics.FromImage(bmp))
        {
            g.CopyFromScreen(clipped.X, clipped.Y, 0, 0, clipped.Size, CopyPixelOperation.SourceCopy);
        }

        byte[] pngBytes;
        using (var ms = new MemoryStream())
        {
            bmp.Save(ms, ImageFormat.Png);
            pngBytes = ms.ToArray();
        }

        return await RecognizePngBytesAsync(pngBytes, timeout, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Runs OCR on in-memory PNG bytes (used by tests and screen capture path).
    /// </summary>
    /// <exception cref="OcrTimeoutException">When recognition exceeds <paramref name="timeout"/>.</exception>
    public async Task<(string Text, double AverageConfidence)> RecognizePngBytesAsync(
        byte[] pngBytes,
        TimeSpan timeout,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(pngBytes);
        if (pngBytes.Length == 0)
            return (string.Empty, 0);

        using var ras = new InMemoryRandomAccessStream();
        var dw = new DataWriter(ras.GetOutputStreamAt(0));
        dw.WriteBytes(pngBytes);
        await dw.StoreAsync().AsTask(cancellationToken).ConfigureAwait(false);
        dw.DetachStream();
        ras.Seek(0);

        var decoder = await BitmapDecoder.CreateAsync(ras).AsTask(cancellationToken).ConfigureAwait(false);
        using SoftwareBitmap softwareBitmap =
            await decoder.GetSoftwareBitmapAsync(BitmapPixelFormat.Bgra8, BitmapAlphaMode.Premultiplied)
                .AsTask(cancellationToken)
                .ConfigureAwait(false);

        OcrEngine? engine = CreateEngine();
        if (engine is null)
            return (string.Empty, 0);

        Task<OcrResult> recognizeTask = engine.RecognizeAsync(softwareBitmap).AsTask(cancellationToken);
        Task timeoutTask = Task.Delay(timeout, cancellationToken);
        Task winner = await Task.WhenAny(recognizeTask, timeoutTask).ConfigureAwait(false);
        if (winner == timeoutTask)
            throw new OcrTimeoutException($"OCR exceeded {timeout.TotalSeconds:0.#}s.");

        OcrResult result = await recognizeTask.ConfigureAwait(false);
        string text = result.Text ?? string.Empty;
        double conf = AverageWordConfidence(result);
        return (text, conf);
    }

    /// <summary>Runs OCR with a 5-second timeout.</summary>
    public async Task<string> ReadTextFromRegionAsync(Rectangle region, CancellationToken cancellationToken = default)
    {
        var (text, _) = await ReadTextFromRegionWithConfidenceAsync(
            region,
            TimeSpan.FromSeconds(5),
            cancellationToken).ConfigureAwait(false);
        return text;
    }

    private OcrEngine? CreateEngine()
    {
        if (_languageTag.Equals("auto", StringComparison.OrdinalIgnoreCase))
            return OcrEngine.TryCreateFromUserProfileLanguages();

        try
        {
            var lang = new Language(_languageTag);
            return OcrEngine.TryCreateFromLanguage(lang);
        }
        catch
        {
            return OcrEngine.TryCreateFromUserProfileLanguages();
        }
    }

    /// <summary>Word-level confidence is not on all SDK projections; use line text density heuristic.</summary>
    private static double AverageWordConfidence(OcrResult result)
    {
        string t = result.Text ?? string.Empty;
        if (t.Length == 0)
            return 0;
        // Heuristic: longer structured OCR output tends to be more reliable for UI gating.
        return Math.Min(1.0, 0.55 + Math.Min(t.Length, 200) / 400.0);
    }
}
