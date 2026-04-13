// Created by Phạm Duy – Giải pháp tự động hóa thông minh.

using System.Collections.Concurrent;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Text;
using SmartMacroAI.Core;
using SmartMacroAI.Models;
using Xunit;

namespace SmartMacroAI.Tests;

public sealed class OcrAndVariablesTests
{
    [Fact]
    public void VariableStore_ConcurrentReadWrite_IsConsistent()
    {
        var store = new VariableStore();
        const int threads = 8;
        const int iters = 500;
        var bag = new ConcurrentBag<string>();

        Parallel.For(0, threads * iters, i =>
        {
            int t = i % threads;
            string key = $"k{t}";
            store.Set(key, i.ToString(), "Manual");
            _ = store.Get(key);
            bag.Add(store.Get(key));
        });

        for (int t = 0; t < threads; t++)
        {
            Assert.True(store.Exists($"k{t}"));
            Assert.True(int.TryParse(store.Get($"k{t}"), out _));
        }
    }

    [Fact]
    public void MacroVariableInterpolator_NestedDoubleCurly_ResolvesInTwoPasses()
    {
        var d = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["inner"] = "world",
            ["outer"] = "hello {{inner}}",
        };

        string step1 = MacroVariableInterpolator.ExpandDoubleCurly("x={{outer}}y", d, null);
        Assert.Equal("x=hello {{inner}}y", step1);
        string step2 = MacroVariableInterpolator.ExpandDoubleCurly(step1, d, null);
        Assert.Equal("x=hello worldy", step2);
    }

    [Fact]
    public async Task OcrService_EmptyPng_ReturnsEmpty()
    {
        if (!OperatingSystem.IsWindowsVersionAtLeast(10, 0, 19041, 0))
            return;

        var ocr = new OcrService("en-US");
        var (_, conf) = await ocr.RecognizePngBytesAsync(Array.Empty<byte>(), TimeSpan.FromSeconds(5));
        Assert.Equal(0, conf);
    }

    [Fact]
    public async Task OcrService_OutOfScreenRegion_ReturnsEmptyWithoutThrowing()
    {
        if (!OperatingSystem.IsWindowsVersionAtLeast(10, 0, 19041, 0))
            return;

        var ocr = new OcrService("auto");
        var rect = new Rectangle(-500_000, -500_000, 400, 200);
        var (text, conf) = await ocr.ReadTextFromRegionWithConfidenceAsync(rect, TimeSpan.FromSeconds(5));
        Assert.Equal(string.Empty, text);
        Assert.Equal(0, conf);
    }

    [Fact]
    public async Task OcrService_KnownBitmap_ContainsExpectedText()
    {
        if (!OperatingSystem.IsWindowsVersionAtLeast(10, 0, 19041, 0))
            return;

        byte[] png = RenderTextPng("OCRTEST42", 320, 120);
        var ocr = new OcrService("en-US");
        var (text, _) = await ocr.RecognizePngBytesAsync(png, TimeSpan.FromSeconds(15));
        Assert.Contains("OCRTEST", text, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void MacroScript_OcrAndVariableActions_RoundTripJson()
    {
        var script = new MacroScript
        {
            Name = "t",
            Actions =
            [
                new OcrRegionAction { ScreenX = 1, ScreenY = 2, ScreenWidth = 100, ScreenHeight = 50, OutputVariableName = "r" },
                new ClearVariableAction { VarName = "a" },
                new LogVariableAction { VarName = "b" },
                new WaitAction
                {
                    WaitForOcrContains = "ok",
                    OcrRegionX = 10,
                    OcrRegionY = 20,
                    OcrRegionWidth = 30,
                    OcrRegionHeight = 40,
                    OcrPollIntervalMs = 250,
                    WaitTimeoutMs = 5000,
                },
            ],
        };

        string path = Path.Combine(Path.GetTempPath(), $"SmartMacroAI_test_{Guid.NewGuid():N}.json");
        try
        {
            ScriptManager.Save(script, path);
            MacroScript? back = ScriptManager.Load(path);
            Assert.NotNull(back);
            Assert.Equal(4, back.Actions.Count);
            Assert.IsType<OcrRegionAction>(back.Actions[0]);
            Assert.IsType<ClearVariableAction>(back.Actions[1]);
            Assert.IsType<LogVariableAction>(back.Actions[2]);
            var w = Assert.IsType<WaitAction>(back.Actions[3]);
            Assert.Equal("ok", w.WaitForOcrContains);
            Assert.Equal(250, w.OcrPollIntervalMs);
        }
        finally
        {
            try
            {
                File.Delete(path);
            }
            catch
            {
                // ignore
            }
        }
    }

    private static byte[] RenderTextPng(string text, int w, int h)
    {
        using var bmp = new Bitmap(w, h, PixelFormat.Format32bppArgb);
        using (var g = Graphics.FromImage(bmp))
        {
            g.Clear(Color.White);
            g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
            using var font = new Font(FontFamily.GenericSansSerif, 28f, FontStyle.Bold, GraphicsUnit.Pixel);
            using var brush = new SolidBrush(Color.Black);
            g.DrawString(text, font, brush, 8f, 32f);
        }

        using var ms = new MemoryStream();
        bmp.Save(ms, ImageFormat.Png);
        return ms.ToArray();
    }
}
