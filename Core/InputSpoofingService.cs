// Created by Phạm Duy – Giải pháp tự động hóa thông minh.

using System.Runtime.InteropServices;

namespace SmartMacroAI.Core;

/// <summary>
/// Low-level <see cref="NativeMethods.SendInput"/> helpers (extra info, wheel, scan-code keys).
/// Fail-safe: errors are swallowed; callers should log if needed.
/// </summary>
public static class InputSpoofingService
{
    public const uint MOUSEEVENTF_WHEEL = 0x0800;

    private static IntPtr ExtraInfo() => NativeMethods.GetMessageExtraInfo();

    public static void SendMouseWheel(int wheelDelta)
    {
        try
        {
            var mi = new NativeMethods.MOUSEINPUT
            {
                dx = 0,
                dy = 0,
                mouseData = unchecked((uint)wheelDelta),
                dwFlags = MOUSEEVENTF_WHEEL,
                time = 0,
                dwExtraInfo = ExtraInfo(),
            };
            var inp = new NativeMethods.INPUT { type = NativeMethods.INPUT_MOUSE, U = new NativeMethods.InputUnion { mi = mi } };
            NativeMethods.SendInput(1, [inp], NativeMethods.SizeOfInput());
        }
        catch
        {
            /* fail-safe */
        }
    }

    private static void SendScan(bool keyUp, ushort scan, bool extended)
    {
        uint f = NativeMethods.KEYEVENTF_SCANCODE | (extended ? NativeMethods.KEYEVENTF_EXTENDEDKEY : 0);
        if (keyUp)
            f |= NativeMethods.KEYEVENTF_KEYUP;

        var ki = new NativeMethods.KEYBDINPUT
        {
            wVk = 0,
            wScan = scan,
            dwFlags = f,
            time = 0,
            dwExtraInfo = ExtraInfo(),
        };
        var inp = new NativeMethods.INPUT { type = NativeMethods.INPUT_KEYBOARD, U = new NativeMethods.InputUnion { ki = ki } };
        NativeMethods.SendInput(1, [inp], NativeMethods.SizeOfInput());
    }

    /// <summary>Key down, hold, key up using scan codes (US layout via <see cref="VkKeyScanW"/>).</summary>
    public static void TapKeyScan(ushort scanCode, bool extended = false)
    {
        try
        {
            SendScan(keyUp: false, scanCode, extended);
            Thread.Sleep(Random.Shared.Next(60, 181));
            SendScan(keyUp: true, scanCode, extended);
        }
        catch
        {
            /* fail-safe */
        }
    }

    /// <summary>Attempts scan-code path including Shift for uppercase / symbols.</summary>
    public static bool TrySendCharHardware(char ch)
    {
        try
        {
            short vkPacked = VkKeyScanW(ch);
            if (vkPacked == -1)
                return false;

            bool needShift = (vkPacked >> 8 & 1) != 0;
            byte vk = unchecked((byte)(vkPacked & 0xFF));
            uint scLo = NativeMethods.MapVirtualKey(vk, NativeMethods.MAPVK_VK_TO_VSC);
            if (scLo == 0)
                return false;

            ushort scan = unchecked((ushort)scLo);
            bool ext = IsExtendedVk(vk);

            if (needShift)
            {
                SendScan(keyUp: false, 0x2A, extended: false);
                Thread.Sleep(Random.Shared.Next(12, 35));
            }

            TapKeyScan(scan, ext);

            if (needShift)
            {
                Thread.Sleep(Random.Shared.Next(12, 35));
                SendScan(keyUp: true, 0x2A, extended: false);
            }

            return true;
        }
        catch
        {
            return false;
        }
    }

    private static bool IsExtendedVk(byte vk) => vk is 0x21 or 0x22 or 0x23 or 0x24 or 0x25 or 0x26 or 0x27 or 0x28
        or 0x2D or 0x2E
        or 0x5B or 0x5C or 0x5D;

    [DllImport("user32.dll")]
    private static extern short VkKeyScanW(char ch);
}
