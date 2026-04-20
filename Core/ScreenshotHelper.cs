// Created by Phạm Duy – Giải pháp tự động hóa thông minh.

using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;

namespace SmartMacroAI.Core;

/// <summary>
/// Utility for capturing window screenshots.
/// Created by Phạm Duy – Giải pháp tự động hóa thông minh.
/// </summary>
public static class ScreenshotHelper
{
    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);

    [DllImport("user32.dll")]
    private static extern IntPtr GetWindowDC(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);

    [DllImport("gdi32.dll")]
    private static extern bool BitBlt(IntPtr hdcDest, int xDest, int yDest,
        int width, int height, IntPtr hdcSrc, int xSrc, int ySrc, uint dwRop);

    [DllImport("gdi32.dll")]
    private static extern IntPtr CreateCompatibleDC(IntPtr hdc);

    [DllImport("gdi32.dll")]
    private static extern IntPtr CreateCompatibleBitmap(IntPtr hdc, int nWidth, int nHeight);

    [DllImport("gdi32.dll")]
    private static extern IntPtr SelectObject(IntPtr hdc, IntPtr hgdiobj);

    [DllImport("gdi32.dll")]
    private static extern bool DeleteObject(IntPtr hObject);

    [DllImport("gdi32.dll")]
    private static extern bool DeleteDC(IntPtr hdc);

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT
    {
        public int Left, Top, Right, Bottom;
    }

    private const uint SRCCOPY = 0x00CC0020;

    /// <summary>
    /// Captures a screenshot of the target window.
    /// Returns saved file path, or null if failed.
    /// </summary>
    public static string? CaptureWindow(IntPtr hwnd, string outputFolder)
    {
        try
        {
            if (hwnd == IntPtr.Zero) return null;

            if (!GetWindowRect(hwnd, out RECT rect))
                return null;

            int width = rect.Right - rect.Left;
            int height = rect.Bottom - rect.Top;

            if (width <= 0 || height <= 0)
                return null;

            Directory.CreateDirectory(outputFolder);

            string fileName = $"error_{DateTime.Now:yyyyMMdd_HHmmss}.png";
            string filePath = Path.Combine(outputFolder, fileName);

            // Use screen DC for BitBlt (more reliable than window DC)
            IntPtr hdcScreen = GetDC(IntPtr.Zero);
            IntPtr hdcMem = CreateCompatibleDC(hdcScreen);
            IntPtr hBitmap = CreateCompatibleBitmap(hdcScreen, width, height);
            IntPtr hOld = SelectObject(hdcMem, hBitmap);

            // Copy from screen coordinates (relative to desktop)
            BitBlt(hdcMem, 0, 0, width, height, hdcScreen,
                   rect.Left, rect.Top, SRCCOPY);

            SelectObject(hdcMem, hOld);

            using var bitmap = Image.FromHbitmap(hBitmap);
            bitmap.Save(filePath, ImageFormat.Png);

            DeleteObject(hBitmap);
            DeleteDC(hdcMem);
            ReleaseDC(IntPtr.Zero, hdcScreen);

            return filePath;
        }
        catch
        {
            return null;
        }
    }

    [DllImport("user32.dll")]
    private static extern IntPtr GetDC(IntPtr hWnd);
}
