// Created by Phạm Duy – Giải pháp tự động hóa thông minh.

using SmartMacroAI.Models;

namespace SmartMacroAI.Core;

public class MacroTemplate
{
    public string Name { get; set; } = "";
    public string Description { get; set; } = "";
    public string Category { get; set; } = "General";
    public List<MacroAction> Actions { get; set; } = [];
    public string TargetWindowTitle { get; set; } = "";
}

public static class MacroTemplateService
{
    public static List<MacroTemplate> GetTemplates() => new()
    {
        new MacroTemplate
        {
            Name = "🔐 Auto Login Website",
            Description = "Tự động đăng nhập website với username/password từ CSV",
            Category = "Web",
            TargetWindowTitle = "",
            Actions = new List<MacroAction>
            {
                new LaunchAndBindAction { DisplayName = "Mở trình duyệt", Url = "{{url}}", Browser = LaunchBrowserKind.Edge, BindTimeoutMs = 30000, PollIntervalMs = 500 },
                new WaitAction { DisplayName = "Chờ trang load", DelayMin = 2000, DelayMax = 3000 },
                new WebClickAction { DisplayName = "Click username", CssSelector = "{{username_selector}}" },
                new WebTypeAction { DisplayName = "Nhập username", CssSelector = "{{username_selector}}", TextToType = "{{username}}" },
                new WebClickAction { DisplayName = "Click password", CssSelector = "{{password_selector}}" },
                new WebTypeAction { DisplayName = "Nhập password", CssSelector = "{{password_selector}}", TextToType = "{{password}}" },
                new WebClickAction { DisplayName = "Click đăng nhập", CssSelector = "{{login_button_selector}}" }
            }
        },
        new MacroTemplate
        {
            Name = "📊 Auto Fill CSV Data",
            Description = "Điền dữ liệu từ file CSV vào form web (lặp nhiều dòng)",
            Category = "Web",
            TargetWindowTitle = "",
            Actions = new List<MacroAction>
            {
                new LaunchAndBindAction { DisplayName = "Mở trang form", Url = "{{form_url}}", Browser = LaunchBrowserKind.Edge, BindTimeoutMs = 30000 },
                new WaitAction { DisplayName = "Chờ form load", DelayMin = 2000, DelayMax = 3000 },
                new RepeatAction
                {
                    DisplayName = "Lặp từng dòng CSV",
                    RepeatCount = 0,
                    IntervalMs = 1000,
                    LoopActions = new List<MacroAction>
                    {
                        new WebClickAction { DisplayName = "Click field 1", CssSelector = "{{field1_selector}}" },
                        new WebTypeAction { DisplayName = "Nhập giá trị 1", CssSelector = "{{field1_selector}}", TextToType = "{{col1}}" },
                        new WebClickAction { DisplayName = "Click field 2", CssSelector = "{{field2_selector}}" },
                        new WebTypeAction { DisplayName = "Nhập giá trị 2", CssSelector = "{{field2_selector}}", TextToType = "{{col2}}" },
                        new WebClickAction { DisplayName = "Click Submit", CssSelector = "{{submit_selector}}" },
                        new WaitAction { DisplayName = "Chờ xử lý", DelayMin = 1000, DelayMax = 2000 }
                    }
                }
            }
        },
        new MacroTemplate
        {
            Name = "🔄 Auto Repeat Task",
            Description = "Lặp lại thao tác click chuột theo chu kỳ",
            Category = "Desktop",
            TargetWindowTitle = "{{target_window}}",
            Actions = new List<MacroAction>
            {
                new RepeatAction
                {
                    DisplayName = "Lặp thao tác",
                    RepeatCount = 10,
                    IntervalMs = 5000,
                    LoopActions = new List<MacroAction>
                    {
                        new ClickAction { DisplayName = "Click vị trí", X = 0, Y = 0 },
                        new WaitAction { DisplayName = "Chờ", DelayMin = 1000, DelayMax = 1500 }
                    }
                }
            }
        },
        new MacroTemplate
        {
            Name = "🔍 Image Detection Loop",
            Description = "Chờ cho đến khi hình ảnh xuất hiện rồi thực hiện hành động",
            Category = "Desktop",
            TargetWindowTitle = "{{target_window}}",
            Actions = new List<MacroAction>
            {
                new RepeatAction
                {
                    DisplayName = "Kiểm tra hình ảnh",
                    RepeatCount = 0,
                    IntervalMs = 2000,
                    LoopActions = new List<MacroAction>
                    {
                        new IfImageAction
                        {
                            DisplayName = "Nếu tìm thấy hình",
                            ImagePath = "{{image_path}}",
                            Threshold = 0.8f,
                            TimeoutMs = 5000,
                            ClickOnFound = true,
                            RandomOffset = 5,
                            ThenActions = new List<MacroAction> { new WaitAction { DisplayName = "Chờ xử lý", DelayMin = 500, DelayMax = 1000 } },
                            ElseActions = new List<MacroAction>()
                        },
                        new WaitAction { DisplayName = "Chờ lại", DelayMin = 500, DelayMax = 800 }
                    }
                }
            }
        },
        new MacroTemplate
        {
            Name = "⌨️ Hotkey Automation",
            Description = "Nhấn tổ hợp phím tắt theo chu kỳ",
            Category = "Desktop",
            TargetWindowTitle = "{{target_window}}",
            Actions = new List<MacroAction>
            {
                new RepeatAction
                {
                    DisplayName = "Lặp phím tắt",
                    RepeatCount = 5,
                    IntervalMs = 3000,
                    LoopActions = new List<MacroAction>
                    {
                        new KeyPressAction { DisplayName = "Nhấn Ctrl+S", KeyName = "S", VirtualKeyCode = 0x53, Modifiers = new KeyModifiers { Ctrl = true }, HoldDurationMs = 100 },
                        new WaitAction { DisplayName = "Chờ", DelayMin = 500, DelayMax = 1000 }
                    }
                }
            }
        },
        new MacroTemplate
        {
            Name = "🚀 Quick Start Blank",
            Description = "Bắt đầu với macro trống - thêm actions thủ công",
            Category = "General",
            TargetWindowTitle = "",
            Actions = new List<MacroAction>()
        }
    };

    public static List<string> GetCategories() =>
        GetTemplates().Select(t => t.Category).Distinct().OrderBy(c => c).ToList();

    public static List<MacroTemplate> GetTemplatesByCategory(string category) =>
        GetTemplates().Where(t => t.Category == category).ToList();
}
