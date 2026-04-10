using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Win32;
using SmartMacroAI.Models;

namespace SmartMacroAI;

public partial class ActionEditDialog : Window
{
    private readonly MacroAction _action;
    private readonly Dictionary<string, TextBox> _fields = [];
    private readonly Dictionary<string, CheckBox> _checkFields = [];

    private static readonly Brush LabelBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#A6ADC8"));
    private static readonly Brush InputBg = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#1E1E2E"));
    private static readonly Brush InputFg = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#CDD6F4"));
    private static readonly Brush InputBorder = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#45475A"));
    private static readonly Brush AccentBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#89B4FA"));

    public ActionEditDialog(MacroAction action)
    {
        InitializeComponent();
        _action = action;
        TxtDialogTitle.Text = $"Edit: {action.DisplayName}";
        BuildFields();
    }

    private void BuildFields()
    {
        switch (_action)
        {
            case ClickAction c:
                AddField("X", c.X.ToString());
                AddField("Y", c.Y.ToString());
                AddCheckField("IsRightClick", c.IsRightClick, "Right-click instead of left-click");
                break;
            case TypeAction t:
                AddField("Text", t.Text);
                AddField("KeyDelayMs", t.KeyDelayMs.ToString());
                break;
            case WaitAction w:
                AddField("Milliseconds", w.Milliseconds.ToString());
                break;
            case IfImageAction img:
                AddField("ImagePath", img.ImagePath, browse: true);
                AddField("Threshold", img.Threshold.ToString("F2"));
                AddCheckField("ClickOnFound", img.ClickOnFound,
                    "Auto-click center of image when found");
                break;
            case IfTextAction txt:
                AddField("Text", txt.Text);
                AddCheckField("IgnoreCase", txt.IgnoreCase, "Case-insensitive comparison");
                AddCheckField("PartialMatch", txt.PartialMatch, "Match if text appears anywhere (substring)");
                break;
            case WebNavigateAction wn:
                AddField("Url", wn.Url);
                break;
            case WebClickAction wc:
                AddField("CssSelector", wc.CssSelector);
                FieldsPanel.Children.Add(new TextBlock
                {
                    Text = "Use CSS (e.g. #id, .class) or Playwright text= / xpath= selectors.",
                    Foreground = LabelBrush,
                    FontSize = 10,
                    TextWrapping = TextWrapping.Wrap,
                    Margin = new Thickness(0, 4, 0, 0),
                });
                break;
            case WebTypeAction wt:
                AddField("CssSelector", wt.CssSelector);
                AddField("TextToType", wt.TextToType);
                FieldsPanel.Children.Add(new TextBlock
                {
                    Text = "Selector: CSS, or xpath=… — field is cleared then filled (FillAsync).",
                    Foreground = LabelBrush,
                    FontSize = 10,
                    TextWrapping = TextWrapping.Wrap,
                    Margin = new Thickness(0, 4, 0, 0),
                });
                break;
        }
    }

    private void AddField(string label, string value, bool browse = false)
    {
        FieldsPanel.Children.Add(new TextBlock
        {
            Text = label.ToUpperInvariant(),
            FontSize = 11,
            FontWeight = FontWeights.SemiBold,
            Foreground = LabelBrush,
            Margin = new Thickness(0, 8, 0, 4),
        });

        var textBox = new TextBox
        {
            Text = value,
            Background = InputBg,
            Foreground = InputFg,
            BorderBrush = InputBorder,
            BorderThickness = new Thickness(1),
            Padding = new Thickness(8, 6, 8, 6),
            FontSize = 12,
            CaretBrush = InputFg,
        };
        _fields[label] = textBox;

        if (browse)
        {
            var panel = new DockPanel();
            var btn = new Button
            {
                Content = "...",
                Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#313244")),
                Foreground = InputFg,
                BorderThickness = new Thickness(0),
                Padding = new Thickness(10, 6, 10, 6),
                Cursor = System.Windows.Input.Cursors.Hand,
                Margin = new Thickness(4, 0, 0, 0),
            };
            btn.Click += (_, _) =>
            {
                var dlg = new OpenFileDialog { Filter = "Image files|*.png;*.jpg;*.jpeg;*.bmp|All|*.*" };
                if (dlg.ShowDialog() == true) textBox.Text = dlg.FileName;
            };
            DockPanel.SetDock(btn, Dock.Right);
            panel.Children.Add(btn);
            panel.Children.Add(textBox);
            FieldsPanel.Children.Add(panel);
        }
        else
        {
            FieldsPanel.Children.Add(textBox);
        }
    }

    private void AddCheckField(string key, bool value, string description)
    {
        var cb = new CheckBox
        {
            IsChecked = value,
            Foreground = InputFg,
            Margin = new Thickness(0, 10, 0, 2),
            Content = new TextBlock
            {
                Text = description,
                Foreground = InputFg,
                FontSize = 12,
                TextWrapping = TextWrapping.Wrap,
            },
        };
        _checkFields[key] = cb;
        FieldsPanel.Children.Add(cb);
    }

    private string GetFieldValue(string key) => _fields.TryGetValue(key, out var tb) ? tb.Text.Trim() : "";
    private bool GetCheckValue(string key) => _checkFields.TryGetValue(key, out var cb) && cb.IsChecked == true;

    private void BtnApply_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            switch (_action)
            {
                case ClickAction c:
                    c.X = int.Parse(GetFieldValue("X"));
                    c.Y = int.Parse(GetFieldValue("Y"));
                    c.IsRightClick = GetCheckValue("IsRightClick");
                    break;
                case TypeAction t:
                    t.Text = GetFieldValue("Text");
                    t.KeyDelayMs = int.Parse(GetFieldValue("KeyDelayMs"));
                    break;
                case WaitAction w:
                    w.Milliseconds = int.Parse(GetFieldValue("Milliseconds"));
                    break;
                case IfImageAction img:
                    img.ImagePath = GetFieldValue("ImagePath");
                    img.Threshold = double.Parse(GetFieldValue("Threshold"));
                    img.ClickOnFound = GetCheckValue("ClickOnFound");
                    break;
                case IfTextAction txt:
                    txt.Text = GetFieldValue("Text");
                    txt.IgnoreCase = GetCheckValue("IgnoreCase");
                    txt.PartialMatch = GetCheckValue("PartialMatch");
                    break;
                case WebNavigateAction wn:
                    wn.Url = GetFieldValue("Url");
                    break;
                case WebClickAction wc:
                    wc.CssSelector = GetFieldValue("CssSelector");
                    break;
                case WebTypeAction wt:
                    wt.CssSelector = GetFieldValue("CssSelector");
                    wt.TextToType = GetFieldValue("TextToType");
                    break;
            }
            DialogResult = true;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Invalid input: {ex.Message}", "Validation Error",
                MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private void BtnCancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
    }
}
