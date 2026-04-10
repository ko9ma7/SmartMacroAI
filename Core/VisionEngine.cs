using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Emgu.CV;
using Emgu.CV.CvEnum;
using TesseractOcr = Tesseract;

namespace SmartMacroAI.Core;

/// <summary>
/// Computer-vision layer for SmartMacroAI.
/// All capture operations use <see cref="Win32Api.CaptureWindow"/> (PrintWindow),
/// which works on background / occluded / minimized windows without bringing them
/// to the foreground.
///
/// • Template matching   → Emgu.CV  (CvInvoke.MatchTemplate)
/// • Text recognition    → Tesseract OCR
/// </summary>
public static class VisionEngine
{
    private static readonly object TessLock = new();
    private static TesseractOcr.TesseractEngine? _tessEngine;

    public static string TessDataPath { get; set; } =
        Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tessdata");

    public static string TessLanguage { get; set; } = "eng";

    // ═══════════════════════════════════════════════════
    //  BITMAP ↔ MAT CONVERSION
    // ═══════════════════════════════════════════════════

    /// <summary>
    /// Converts a <see cref="Bitmap"/> to an Emgu.CV <see cref="Mat"/>
    /// by encoding to PNG in memory and decoding with OpenCV.
    /// Works with any Emgu.CV version — no extension-method dependency.
    /// </summary>
    private static Mat BitmapToMat(Bitmap bitmap)
    {
        using var ms = new MemoryStream();
        bitmap.Save(ms, ImageFormat.Bmp);
        byte[] bytes = ms.ToArray();
        var mat = new Mat();
        CvInvoke.Imdecode(bytes, ImreadModes.ColorBgr, mat);
        return mat;
    }

    // ═══════════════════════════════════════════════════
    //  BACKGROUND WINDOW CAPTURE
    // ═══════════════════════════════════════════════════

    /// <summary>
    /// Captures a background window into a <see cref="Bitmap"/> using
    /// <c>PrintWindow</c> + <c>PW_RENDERFULLCONTENT</c>.
    /// The physical mouse/keyboard are never touched and the window
    /// does NOT need to be in the foreground.
    /// </summary>
    public static Bitmap CaptureHiddenWindow(IntPtr hwnd)
    {
        Bitmap? bmp = Win32Api.CaptureWindow(hwnd);
        if (bmp is null)
            throw new InvalidOperationException(
                $"Failed to capture window (HWND=0x{hwnd:X}). " +
                "The handle may be invalid or the window may have zero size.");
        return bmp;
    }

    // ═══════════════════════════════════════════════════
    //  TEMPLATE MATCHING  (Emgu.CV / OpenCV)
    // ═══════════════════════════════════════════════════

    /// <summary>
    /// Captures the target window in the background and searches for the
    /// template image inside it using OpenCV template matching.
    /// Returns the centre-point (client-relative) of the best match,
    /// or <c>null</c> when the confidence is below <paramref name="threshold"/>.
    /// </summary>
    public static Point? FindImageOnWindow(IntPtr hwnd, string templatePath, double threshold = 0.8)
    {
        if (!File.Exists(templatePath))
            throw new FileNotFoundException("Template image not found.", templatePath);

        using Bitmap captured = CaptureHiddenWindow(hwnd);
        return FindImageInBitmap(captured, templatePath, threshold);
    }

    /// <summary>
    /// Runs template matching on a pre-captured bitmap.
    /// Useful when you want to capture once and run multiple searches.
    /// </summary>
    public static Point? FindImageInBitmap(Bitmap source, string templatePath, double threshold = 0.8)
    {
        using Mat sourceMat   = BitmapToMat(source);
        using Mat templateMat = CvInvoke.Imread(templatePath, ImreadModes.ColorBgr);

        if (templateMat.IsEmpty)
            throw new InvalidOperationException($"Failed to load template image: {templatePath}");

        if (templateMat.Width > sourceMat.Width || templateMat.Height > sourceMat.Height)
            return null;

        using var result = new Mat();
        CvInvoke.MatchTemplate(sourceMat, templateMat, result, TemplateMatchingType.CcoeffNormed);

        double minVal = 0, maxVal = 0;
        Point minLoc = default, maxLoc = default;
        CvInvoke.MinMaxLoc(result, ref minVal, ref maxVal, ref minLoc, ref maxLoc);

        if (maxVal < threshold)
            return null;

        int centerX = maxLoc.X + templateMat.Width  / 2;
        int centerY = maxLoc.Y + templateMat.Height / 2;
        return new Point(centerX, centerY);
    }

    /// <summary>
    /// Returns the raw match confidence (0.0 – 1.0) for diagnostic / UI display.
    /// </summary>
    public static (Point Location, double Confidence)? FindImageOnWindowDetailed(
        IntPtr hwnd, string templatePath)
    {
        if (!File.Exists(templatePath))
            throw new FileNotFoundException("Template image not found.", templatePath);

        using Bitmap captured = CaptureHiddenWindow(hwnd);
        using Mat sourceMat   = BitmapToMat(captured);
        using Mat templateMat = CvInvoke.Imread(templatePath, ImreadModes.ColorBgr);

        if (templateMat.IsEmpty)
            return null;

        if (templateMat.Width > sourceMat.Width || templateMat.Height > sourceMat.Height)
            return null;

        using var result = new Mat();
        CvInvoke.MatchTemplate(sourceMat, templateMat, result, TemplateMatchingType.CcoeffNormed);

        double minVal = 0, maxVal = 0;
        Point minLoc = default, maxLoc = default;
        CvInvoke.MinMaxLoc(result, ref minVal, ref maxVal, ref minLoc, ref maxLoc);

        var center = new Point(
            maxLoc.X + templateMat.Width  / 2,
            maxLoc.Y + templateMat.Height / 2);

        return (center, maxVal);
    }

    // ═══════════════════════════════════════════════════
    //  OCR  (Tesseract)
    // ═══════════════════════════════════════════════════

    /// <summary>
    /// Captures the target window in the background and runs Tesseract OCR
    /// to extract all visible text.
    /// Requires tessdata/{lang}.traineddata to be present.
    /// </summary>
    public static string ExtractTextFromWindow(IntPtr hwnd)
    {
        using Bitmap captured = CaptureHiddenWindow(hwnd);
        return ExtractTextFromBitmap(captured);
    }

    /// <summary>
    /// Runs Tesseract OCR on a pre-captured bitmap.
    /// </summary>
    public static string ExtractTextFromBitmap(Bitmap bitmap)
    {
        var engine = GetTesseractEngine();
        if (engine is null)
            return "[OCR unavailable — tessdata not found. " +
                   $"Place {TessLanguage}.traineddata in: {TessDataPath}]";

        using var ms = new MemoryStream();
        bitmap.Save(ms, ImageFormat.Png);
        using var pix  = TesseractOcr.Pix.LoadFromMemory(ms.ToArray());
        using var page = engine.Process(pix);
        return page.GetText().Trim();
    }

    /// <summary>
    /// Checks whether the required tessdata files exist and the engine can be initialised.
    /// </summary>
    public static bool IsTesseractAvailable()
    {
        string trainedDataFile = Path.Combine(TessDataPath, $"{TessLanguage}.traineddata");
        return File.Exists(trainedDataFile);
    }

    private static TesseractOcr.TesseractEngine? GetTesseractEngine()
    {
        if (_tessEngine is not null) return _tessEngine;

        lock (TessLock)
        {
            if (_tessEngine is not null) return _tessEngine;

            if (!IsTesseractAvailable())
                return null;

            _tessEngine = new TesseractOcr.TesseractEngine(
                TessDataPath,
                TessLanguage,
                TesseractOcr.EngineMode.Default);

            return _tessEngine;
        }
    }

    /// <summary>
    /// Releases the cached Tesseract engine (call on app shutdown).
    /// </summary>
    public static void Shutdown()
    {
        lock (TessLock)
        {
            _tessEngine?.Dispose();
            _tessEngine = null;
        }
    }
}
