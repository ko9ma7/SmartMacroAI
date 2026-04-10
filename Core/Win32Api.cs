using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Text;

namespace SmartMacroAI.Core;

/// <summary>
/// Central Win32 interop layer for SmartMacroAI.
/// Every automation call goes through here — all methods target window handles
/// directly so the physical mouse and keyboard are NEVER hijacked.
///
/// Created by Phạm Duy - Giải pháp tự động hóa thông minh.
/// </summary>
public static class Win32Api
{
    // ═══════════════════════════════════════════════
    //  CONSTANTS — Window Messages
    // ═══════════════════════════════════════════════

    public const uint WM_LBUTTONDOWN   = 0x0201;
    public const uint WM_LBUTTONUP     = 0x0202;
    public const uint WM_RBUTTONDOWN   = 0x0204;
    public const uint WM_RBUTTONUP     = 0x0205;
    public const uint WM_MBUTTONDOWN   = 0x0207;
    public const uint WM_MBUTTONUP     = 0x0208;
    public const uint WM_KEYDOWN       = 0x0100;
    public const uint WM_KEYUP         = 0x0101;
    public const uint WM_CHAR          = 0x0102;
    public const uint WM_SETTEXT       = 0x000C;
    public const uint WM_GETTEXT       = 0x000D;
    public const uint WM_GETTEXTLENGTH = 0x000E;
    public const uint WM_MOUSEMOVE     = 0x0200;
    public const uint WM_CLOSE         = 0x0010;

    public const uint MK_LBUTTON = 0x0001;
    public const uint MK_RBUTTON = 0x0002;

    // ═══════════════════════════════════════════════
    //  CONSTANTS — ShowWindow / Hotkey
    // ═══════════════════════════════════════════════

    public const int SW_HIDE    = 0;
    public const int SW_SHOW    = 5;
    public const int SW_RESTORE = 9;
    public const int WM_HOTKEY  = 0x0312;

    // ═══════════════════════════════════════════════
    //  CONSTANTS — PrintWindow / GDI
    // ═══════════════════════════════════════════════

    public const uint PW_CLIENTONLY         = 0x00000001;
    public const uint PW_RENDERFULLCONTENT  = 0x00000002;
    public const uint SRCCOPY              = 0x00CC0020;

    // ═══════════════════════════════════════════════
    //  P/INVOKE — Message Dispatch
    // ═══════════════════════════════════════════════

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern IntPtr SendMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern IntPtr SendMessage(IntPtr hWnd, uint msg, IntPtr wParam, StringBuilder lParam);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

    // ═══════════════════════════════════════════════
    //  P/INVOKE — Window Discovery
    // ═══════════════════════════════════════════════

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern IntPtr FindWindow(string? lpClassName, string? lpWindowName);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern IntPtr FindWindowEx(IntPtr hWndParent, IntPtr hWndChildAfter,
                                              string? lpszClass, string? lpszWindow);

    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern int GetWindowTextLength(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool IsWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool IsIconic(IntPtr hWnd);

    // ═══════════════════════════════════════════════
    //  P/INVOKE — Window Geometry
    // ═══════════════════════════════════════════════

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    // ═══════════════════════════════════════════════
    //  P/INVOKE — Background Window Capture
    // ═══════════════════════════════════════════════

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool PrintWindow(IntPtr hWnd, IntPtr hdcBlt, uint nFlags);

    // ═══════════════════════════════════════════════
    //  P/INVOKE — GDI (for bitmap operations)
    // ═══════════════════════════════════════════════

    [DllImport("user32.dll")]
    public static extern IntPtr GetDC(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);

    [DllImport("gdi32.dll")]
    public static extern IntPtr CreateCompatibleDC(IntPtr hdc);

    [DllImport("gdi32.dll")]
    public static extern IntPtr CreateCompatibleBitmap(IntPtr hdc, int nWidth, int nHeight);

    [DllImport("gdi32.dll")]
    public static extern IntPtr SelectObject(IntPtr hdc, IntPtr hgdiobj);

    [DllImport("gdi32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool BitBlt(IntPtr hdcDest, int xDest, int yDest, int wDest, int hDest,
                                      IntPtr hdcSrc, int xSrc, int ySrc, uint rop);

    [DllImport("gdi32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool DeleteObject(IntPtr hObject);

    [DllImport("gdi32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool DeleteDC(IntPtr hdc);

    // ═══════════════════════════════════════════════
    //  P/INVOKE — Process Info
    // ═══════════════════════════════════════════════

    [DllImport("user32.dll", SetLastError = true)]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    // ═══════════════════════════════════════════════
    //  P/INVOKE — Coordinate Conversion
    // ═══════════════════════════════════════════════

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool ScreenToClient(IntPtr hWnd, ref POINT lpPoint);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool ClientToScreen(IntPtr hWnd, ref POINT lpPoint);

    // ═══════════════════════════════════════════════
    //  P/INVOKE — Window Visibility & Hotkeys
    // ═══════════════════════════════════════════════

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    // ═══════════════════════════════════════════════
    //  STRUCTURES
    // ═══════════════════════════════════════════════

    [StructLayout(LayoutKind.Sequential)]
    public struct POINT
    {
        public int X;
        public int Y;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;

        public int Width  => Right - Left;
        public int Height => Bottom - Top;
    }

    // ═══════════════════════════════════════════════
    //  HELPER — Coordinate Packing
    // ═══════════════════════════════════════════════

    public static IntPtr MakeLParam(int x, int y)
        => (IntPtr)((y << 16) | (x & 0xFFFF));

    // ═══════════════════════════════════════════════
    //  STEALTH CLICK — Humanized async sequence that
    //  sends MOUSEMOVE → BUTTONDOWN → delay → BUTTONUP
    //  to defeat apps that reject instant PostMessage clicks.
    // ═══════════════════════════════════════════════

    public static async Task ControlClickAsync(IntPtr hWnd, int x, int y)
    {
        IntPtr lParam = MakeLParam(x, y);
        PostMessage(hWnd, WM_MOUSEMOVE, IntPtr.Zero, lParam);
        await Task.Delay(10);
        PostMessage(hWnd, WM_LBUTTONDOWN, (IntPtr)MK_LBUTTON, lParam);
        await Task.Delay(Random.Shared.Next(20, 50));
        PostMessage(hWnd, WM_LBUTTONUP, IntPtr.Zero, lParam);
    }

    public static async Task ControlRightClickAsync(IntPtr hWnd, int x, int y)
    {
        IntPtr lParam = MakeLParam(x, y);
        PostMessage(hWnd, WM_MOUSEMOVE, IntPtr.Zero, lParam);
        await Task.Delay(10);
        PostMessage(hWnd, WM_RBUTTONDOWN, (IntPtr)MK_RBUTTON, lParam);
        await Task.Delay(Random.Shared.Next(20, 50));
        PostMessage(hWnd, WM_RBUTTONUP, IntPtr.Zero, lParam);
    }

    // ═══════════════════════════════════════════════
    //  STEALTH KEYBOARD — PostMessage-based
    // ═══════════════════════════════════════════════

    public static void ControlSendText(IntPtr hWnd, string text)
    {
        foreach (char c in text)
            PostMessage(hWnd, WM_CHAR, (IntPtr)c, IntPtr.Zero);
    }

    public static void ControlSendKey(IntPtr hWnd, int virtualKeyCode)
    {
        PostMessage(hWnd, WM_KEYDOWN, (IntPtr)virtualKeyCode, IntPtr.Zero);
        PostMessage(hWnd, WM_KEYUP, (IntPtr)virtualKeyCode, IntPtr.Zero);
    }

    /// <summary>
    /// Shows or hides a window without affecting its message queue.
    /// PostMessage-based automation continues to work on hidden windows.
    /// </summary>
    public static void SetWindowVisibility(IntPtr hwnd, bool visible)
    {
        if (hwnd != IntPtr.Zero && IsWindow(hwnd))
            ShowWindow(hwnd, visible ? SW_SHOW : SW_HIDE);
    }

    // ═══════════════════════════════════════════════
    //  WINDOW INFO HELPERS
    // ═══════════════════════════════════════════════

    public static string GetWindowTitle(IntPtr hWnd)
    {
        int length = GetWindowTextLength(hWnd);
        if (length == 0) return string.Empty;

        var sb = new StringBuilder(length + 1);
        GetWindowText(hWnd, sb, sb.Capacity);
        return sb.ToString();
    }

    public static string GetWindowClassName(IntPtr hWnd)
    {
        var sb = new StringBuilder(256);
        GetClassName(hWnd, sb, sb.Capacity);
        return sb.ToString();
    }

    /// <summary>
    /// Returns all visible top-level windows with non-empty titles.
    /// Used by the UI to populate target-window pickers.
    /// </summary>
    public static List<(IntPtr Handle, string Title)> GetAllVisibleWindows()
    {
        var windows = new List<(IntPtr, string)>();

        EnumWindows((hWnd, _) =>
        {
            if (IsWindowVisible(hWnd))
            {
                string title = GetWindowTitle(hWnd);
                if (!string.IsNullOrWhiteSpace(title))
                    windows.Add((hWnd, title));
            }
            return true;
        }, IntPtr.Zero);

        return windows;
    }

    /// <summary>
    /// Finds the first visible top-level window whose title contains
    /// <paramref name="partialTitle"/> (case-insensitive).
    /// Returns <see cref="IntPtr.Zero"/> if no match.
    /// </summary>
    public static IntPtr FindWindowByPartialTitle(string partialTitle)
    {
        IntPtr found = IntPtr.Zero;

        EnumWindows((hWnd, _) =>
        {
            if (!IsWindowVisible(hWnd)) return true;

            string title = GetWindowTitle(hWnd);
            if (title.Contains(partialTitle, StringComparison.OrdinalIgnoreCase))
            {
                found = hWnd;
                return false; // stop enumeration
            }
            return true;
        }, IntPtr.Zero);

        return found;
    }

    // ═══════════════════════════════════════════════
    //  BACKGROUND WINDOW CAPTURE (PrintWindow)
    //  ── works even if the window is behind other
    //     windows, minimized, or off-screen.
    // ═══════════════════════════════════════════════

    /// <summary>
    /// Captures a screenshot of the window identified by <paramref name="hWnd"/>
    /// using <c>PrintWindow</c> with <c>PW_RENDERFULLCONTENT</c>.
    /// The target window does NOT need to be in the foreground.
    /// Returns null when the handle is invalid or the window has zero size.
    /// </summary>
    public static Bitmap? CaptureWindow(IntPtr hWnd)
    {
        if (hWnd == IntPtr.Zero || !IsWindow(hWnd))
            return null;

        if (!GetWindowRect(hWnd, out RECT rect))
            return null;

        int width  = rect.Width;
        int height = rect.Height;
        if (width <= 0 || height <= 0)
            return null;

        var bmp = new Bitmap(width, height, PixelFormat.Format32bppArgb);
        using Graphics gfx = Graphics.FromImage(bmp);
        IntPtr hdc = gfx.GetHdc();

        try
        {
            // PW_RENDERFULLCONTENT (0x2) captures DWM-composed content
            // even when the window is occluded or minimized.
            bool ok = PrintWindow(hWnd, hdc, PW_RENDERFULLCONTENT);
            if (!ok)
            {
                // Fallback: try client-only capture
                ok = PrintWindow(hWnd, hdc, PW_CLIENTONLY);
            }

            return ok ? bmp : null;
        }
        finally
        {
            gfx.ReleaseHdc(hdc);
        }
    }

    /// <summary>
    /// Same as <see cref="CaptureWindow"/> but captures only the client area
    /// (excludes title bar and borders).
    /// </summary>
    public static Bitmap? CaptureWindowClient(IntPtr hWnd)
    {
        if (hWnd == IntPtr.Zero || !IsWindow(hWnd))
            return null;

        if (!GetClientRect(hWnd, out RECT rect))
            return null;

        int width  = rect.Width;
        int height = rect.Height;
        if (width <= 0 || height <= 0)
            return null;

        var bmp = new Bitmap(width, height, PixelFormat.Format32bppArgb);
        using Graphics gfx = Graphics.FromImage(bmp);
        IntPtr hdc = gfx.GetHdc();

        try
        {
            bool ok = PrintWindow(hWnd, hdc, PW_RENDERFULLCONTENT | PW_CLIENTONLY);
            return ok ? bmp : null;
        }
        finally
        {
            gfx.ReleaseHdc(hdc);
        }
    }

    /// <summary>
    /// Alternative capture path using BitBlt from the window's own DC.
    /// Useful as a fallback when PrintWindow returns a black frame
    /// (some hardware-accelerated apps).
    /// </summary>
    public static Bitmap? CaptureWindowBitBlt(IntPtr hWnd)
    {
        if (hWnd == IntPtr.Zero || !IsWindow(hWnd))
            return null;

        if (!GetWindowRect(hWnd, out RECT rect))
            return null;

        int width  = rect.Width;
        int height = rect.Height;
        if (width <= 0 || height <= 0)
            return null;

        IntPtr hdcWindow = GetDC(hWnd);
        IntPtr hdcMemory = CreateCompatibleDC(hdcWindow);
        IntPtr hBitmap   = CreateCompatibleBitmap(hdcWindow, width, height);
        IntPtr hOld      = SelectObject(hdcMemory, hBitmap);

        try
        {
            BitBlt(hdcMemory, 0, 0, width, height, hdcWindow, 0, 0, SRCCOPY);
            SelectObject(hdcMemory, hOld);

            return Image.FromHbitmap(hBitmap);
        }
        finally
        {
            DeleteObject(hBitmap);
            DeleteDC(hdcMemory);
            ReleaseDC(hWnd, hdcWindow);
        }
    }
}
