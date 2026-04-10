using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using Microsoft.Win32;
using SmartMacroAI.Core;
using SmartMacroAI.Models;
using SmartMacroAI.ViewModels;
using WinForms = System.Windows.Forms;
using Drawing = System.Drawing;

namespace SmartMacroAI;

public partial class MainWindow : Window
{
    private MacroScript _currentScript = new();
    private readonly ObservableCollection<MacroAction> _actions = [];
    private MacroEngine? _macroEngine;
    private CancellationTokenSource? _cts;
    private MacroRecorder? _recorder;
    private int _runsToday;

    private readonly ObservableCollection<DashboardRowVm> _dashboardRows = [];

    // ── Hotkey & Tray ──
    private const int HOTKEY_TOGGLE_APP    = 1;
    private const int HOTKEY_TOGGLE_TARGET = 2;
    private HotkeySettings _hotkeySettings = new();
    private HwndSource? _hwndSource;
    private WinForms.NotifyIcon? _trayIcon;
    private bool _appHidden;

    // ── Stealth Tracker (HWND → title) ──
    private readonly Dictionary<IntPtr, string> _hiddenWindows = new();
    private readonly ObservableCollection<StealthWindowVm> _stealthRows = [];

    public MainWindow()
    {
        InitializeComponent();
        DashboardGrid.ItemsSource = _dashboardRows;
        StealthGrid.ItemsSource = _stealthRows;
        _hotkeySettings = HotkeySettings.Load();
        InitializeTrayIcon();
        SyncScriptToUi();
        LoadDashboard();
    }

    // ═══════════════════════════════════════════════════
    //  WINDOW LOADED
    // ═══════════════════════════════════════════════════

    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
        _hwndSource = HwndSource.FromHwnd(new WindowInteropHelper(this).Handle);
        _hwndSource?.AddHook(WndProc);
        RegisterHotkeys();
        InitSettingsUi();
    }

    // ═══════════════════════════════════════════════════
    //  GLOBAL HOTKEYS
    // ═══════════════════════════════════════════════════

    private void RegisterHotkeys()
    {
        IntPtr hwnd = new WindowInteropHelper(this).Handle;
        Win32Api.RegisterHotKey(hwnd, HOTKEY_TOGGLE_APP,
            (uint)_hotkeySettings.ToggleAppModifier, (uint)_hotkeySettings.ToggleAppKey);
        Win32Api.RegisterHotKey(hwnd, HOTKEY_TOGGLE_TARGET,
            (uint)_hotkeySettings.ToggleTargetModifier, (uint)_hotkeySettings.ToggleTargetKey);
        AppendLog($"Hotkeys registered: App={_hotkeySettings.ToggleAppDisplay}, Target={_hotkeySettings.ToggleTargetDisplay}");
    }

    private void UnregisterHotkeys()
    {
        IntPtr hwnd = new WindowInteropHelper(this).Handle;
        if (hwnd == IntPtr.Zero) return;
        Win32Api.UnregisterHotKey(hwnd, HOTKEY_TOGGLE_APP);
        Win32Api.UnregisterHotKey(hwnd, HOTKEY_TOGGLE_TARGET);
    }

    private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        if (msg == Win32Api.WM_HOTKEY)
        {
            int id = wParam.ToInt32();
            if (id == HOTKEY_TOGGLE_APP) ToggleAppVisibility();
            else if (id == HOTKEY_TOGGLE_TARGET) ToggleTargetVisibility();
            handled = true;
        }
        return IntPtr.Zero;
    }

    private void ToggleAppVisibility()
    {
        if (_appHidden)
        {
            Show();
            WindowState = WindowState.Normal;
            ShowInTaskbar = true;
            Win32Api.SetForegroundWindow(new WindowInteropHelper(this).Handle);
            _appHidden = false;
        }
        else
        {
            Hide();
            ShowInTaskbar = false;
            _appHidden = true;
        }
    }

    private void ToggleTargetVisibility()
    {
        string targetTitle = CmbTargetWindow.Text.Trim();
        if (string.IsNullOrWhiteSpace(targetTitle)) return;

        var alreadyHidden = _hiddenWindows
            .FirstOrDefault(kv => kv.Value.Contains(targetTitle, StringComparison.OrdinalIgnoreCase));

        if (alreadyHidden.Key != IntPtr.Zero)
        {
            StealthShowWindow(alreadyHidden.Key);
            Dispatcher.Invoke(() => AppendLog($"Hotkey: restored \"{alreadyHidden.Value}\""));
        }
        else
        {
            IntPtr hwnd = Win32Api.FindWindowByPartialTitle(targetTitle);
            if (hwnd != IntPtr.Zero)
            {
                string title = Win32Api.GetWindowTitle(hwnd);
                StealthHideWindow(hwnd, title);
                Dispatcher.Invoke(() => AppendLog($"Hotkey: hidden \"{title}\""));
            }
        }
    }

    // ═══════════════════════════════════════════════════
    //  STEALTH TRACKER — central hide/show + tray sync
    // ═══════════════════════════════════════════════════

    private void StealthHideWindow(IntPtr hwnd, string title)
    {
        if (_hiddenWindows.ContainsKey(hwnd)) return;
        Win32Api.SetWindowVisibility(hwnd, false);
        _hiddenWindows[hwnd] = title;
        SyncStealthRowState(hwnd, true);
        RebuildTrayMenu();
    }

    private void StealthShowWindow(IntPtr hwnd)
    {
        Win32Api.SetWindowVisibility(hwnd, true);
        _hiddenWindows.Remove(hwnd);
        SyncStealthRowState(hwnd, false);
        RebuildTrayMenu();
    }

    private void SyncStealthRowState(IntPtr hwnd, bool hidden)
    {
        var row = _stealthRows.FirstOrDefault(r => r.Hwnd == hwnd);
        if (row is not null) row.IsHidden = hidden;
    }

    private void ShowAllHiddenWindows()
    {
        int count = 0;
        foreach (var (hwnd, _) in _hiddenWindows.ToList())
        {
            if (Win32Api.IsWindow(hwnd))
            {
                Win32Api.SetWindowVisibility(hwnd, true);
                count++;
            }
        }
        _hiddenWindows.Clear();

        foreach (var row in _dashboardRows)
        {
            if (row.TargetHwnd != IntPtr.Zero && row.StealthMode)
            {
                Win32Api.SetWindowVisibility(row.TargetHwnd, true);
                count++;
            }
        }

        foreach (var sr in _stealthRows) sr.IsHidden = false;
        RebuildTrayMenu();

        if (_appHidden)
        {
            Show();
            WindowState = WindowState.Normal;
            ShowInTaskbar = true;
            _appHidden = false;
        }

        AppendLog($"Emergency: restored {count} hidden window(s) + SmartMacroAI.");
    }

    // ═══════════════════════════════════════════════════
    //  STEALTH MANAGER UI
    // ═══════════════════════════════════════════════════

    private void LoadStealthManager()
    {
        _stealthRows.Clear();
        IntPtr myHwnd = new WindowInteropHelper(this).Handle;

        foreach (var (hwnd, title) in _hiddenWindows)
        {
            _stealthRows.Add(new StealthWindowVm { Hwnd = hwnd, WindowTitle = title, IsHidden = true });
        }

        foreach (var (hwnd, title) in Win32Api.GetAllVisibleWindows())
        {
            if (hwnd == myHwnd) continue;
            if (_hiddenWindows.ContainsKey(hwnd)) continue;
            _stealthRows.Add(new StealthWindowVm { Hwnd = hwnd, WindowTitle = title, IsHidden = false });
        }
    }

    private void StealthToggle_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button { DataContext: StealthWindowVm row }) return;

        if (row.IsHidden)
        {
            StealthShowWindow(row.Hwnd);
            AppendLog($"[Stealth] Restored: \"{row.WindowTitle}\"");
        }
        else
        {
            StealthHideWindow(row.Hwnd, row.WindowTitle);
            AppendLog($"[Stealth] Hidden: \"{row.WindowTitle}\"");
        }
    }

    private void BtnRefreshStealth_Click(object sender, RoutedEventArgs e) => LoadStealthManager();

    private void BtnShowAllStealth_Click(object sender, RoutedEventArgs e) => ShowAllHiddenWindows();

    // ═══════════════════════════════════════════════════
    //  SYSTEM TRAY — dynamic menu
    // ═══════════════════════════════════════════════════

    private void InitializeTrayIcon()
    {
        _trayIcon = new WinForms.NotifyIcon
        {
            Icon = CreateTrayIcon(),
            Text = "SmartMacroAI — Phạm Duy",
            Visible = true,
        };
        _trayIcon.DoubleClick += (_, _) =>
            Dispatcher.Invoke(() => { Show(); WindowState = WindowState.Normal; ShowInTaskbar = true; _appHidden = false;
                Win32Api.SetForegroundWindow(new WindowInteropHelper(this).Handle); });
        RebuildTrayMenu();
    }

    private void RebuildTrayMenu()
    {
        if (_trayIcon is null) return;

        var menu = new WinForms.ContextMenuStrip();
        menu.Items.Add("Show SmartMacroAI", null, (_, _) =>
            Dispatcher.Invoke(() => { Show(); WindowState = WindowState.Normal; ShowInTaskbar = true; _appHidden = false;
                Win32Api.SetForegroundWindow(new WindowInteropHelper(this).Handle); }));
        menu.Items.Add(new WinForms.ToolStripSeparator());

        if (_hiddenWindows.Count > 0)
        {
            var sub = new WinForms.ToolStripMenuItem($"Hidden Windows ({_hiddenWindows.Count})");
            foreach (var (hwnd, title) in _hiddenWindows.ToList())
            {
                var capturedHwnd = hwnd;
                string shortTitle = title.Length > 40 ? string.Concat(title.AsSpan(0, 37), "...") : title;
                sub.DropDownItems.Add(shortTitle, null, (_, _) =>
                    Dispatcher.Invoke(() =>
                    {
                        StealthShowWindow(capturedHwnd);
                        AppendLog($"[Tray] Restored: \"{title}\"");
                    }));
            }
            menu.Items.Add(sub);
        }
        else
        {
            var noItems = new WinForms.ToolStripMenuItem("(no hidden windows)") { Enabled = false };
            menu.Items.Add(noItems);
        }

        menu.Items.Add("Show All Hidden Windows", null, (_, _) =>
            Dispatcher.Invoke(ShowAllHiddenWindows));
        menu.Items.Add("Stop All Macros", null, (_, _) =>
            Dispatcher.Invoke(() => BtnStopAllMacros_Click(this, new RoutedEventArgs())));
        menu.Items.Add(new WinForms.ToolStripSeparator());
        menu.Items.Add("Exit", null, (_, _) =>
            Dispatcher.Invoke(() => { ShowAllHiddenWindows(); Close(); }));

        _trayIcon.ContextMenuStrip = menu;
    }

    private static Drawing.Icon CreateTrayIcon()
    {
        var bmp = new Drawing.Bitmap(16, 16);
        using var g = Drawing.Graphics.FromImage(bmp);
        g.Clear(Drawing.Color.FromArgb(137, 180, 250));
        using var font = new Drawing.Font("Segoe UI", 9, Drawing.FontStyle.Bold);
        g.DrawString("S", font, Drawing.Brushes.Black, 1, 0);
        return Drawing.Icon.FromHandle(bmp.GetHicon());
    }

    // ═══════════════════════════════════════════════════
    //  NAVIGATION
    // ═══════════════════════════════════════════════════

    private void SetActiveView(string viewName)
    {
        DashboardView.Visibility = Visibility.Collapsed;
        MacroEditorView.Visibility = Visibility.Collapsed;
        ImageRecognitionView.Visibility = Visibility.Collapsed;
        OcrEngineView.Visibility = Visibility.Collapsed;
        StealthManagerView.Visibility = Visibility.Collapsed;
        SettingsView.Visibility = Visibility.Collapsed;
        AboutView.Visibility = Visibility.Collapsed;
        ResetSidebarButtons();

        switch (viewName)
        {
            case "Dashboard":
                DashboardView.Visibility = Visibility.Visible;
                BtnDashboard.Style = (Style)FindResource("SidebarButtonActiveStyle");
                TxtPageTitle.Text = "Dashboard";
                TxtPageSubtitle.Text = "Multi-tasking hub — run multiple macros simultaneously";
                LoadDashboard();
                break;
            case "MacroEditor":
                MacroEditorView.Visibility = Visibility.Visible;
                BtnMacroEditor.Style = (Style)FindResource("SidebarButtonActiveStyle");
                TxtPageTitle.Text = "Macro Editor";
                TxtPageSubtitle.Text = "Design your automation workflow with drag & drop";
                break;
            case "ImageRecognition":
                ImageRecognitionView.Visibility = Visibility.Visible;
                BtnImageRecognition.Style = (Style)FindResource("SidebarButtonActiveStyle");
                TxtPageTitle.Text = "Image Recognition";
                TxtPageSubtitle.Text = "Configure template matching for background windows";
                break;
            case "OcrEngine":
                OcrEngineView.Visibility = Visibility.Visible;
                BtnOcrEngine.Style = (Style)FindResource("SidebarButtonActiveStyle");
                TxtPageTitle.Text = "OCR Engine";
                TxtPageSubtitle.Text = "Text recognition settings and testing";
                break;
            case "StealthManager":
                StealthManagerView.Visibility = Visibility.Visible;
                BtnStealthManager.Style = (Style)FindResource("SidebarButtonActiveStyle");
                TxtPageTitle.Text = "Stealth Manager";
                TxtPageSubtitle.Text = "Ẩn/hiện cửa sổ — PostMessage vẫn hoạt động trên cửa sổ ẩn";
                LoadStealthManager();
                break;
            case "Settings":
                SettingsView.Visibility = Visibility.Visible;
                BtnSettingsNav.Style = (Style)FindResource("SidebarButtonActiveStyle");
                TxtPageTitle.Text = "Settings";
                TxtPageSubtitle.Text = "Global hotkeys and stealth configuration";
                break;
            case "About":
                AboutView.Visibility = Visibility.Visible;
                BtnAbout.Style = (Style)FindResource("SidebarButtonActiveStyle");
                TxtPageTitle.Text = "About";
                TxtPageSubtitle.Text = "SmartMacroAI — Created by Phạm Duy";
                break;
        }
    }

    private void ResetSidebarButtons()
    {
        var s = (Style)FindResource("SidebarButtonStyle");
        BtnDashboard.Style = s;
        BtnMacroEditor.Style = s;
        BtnImageRecognition.Style = s;
        BtnOcrEngine.Style = s;
        BtnStealthManager.Style = s;
        BtnSettingsNav.Style = s;
        BtnAbout.Style = s;
    }

    private void BtnDashboard_Click(object sender, RoutedEventArgs e) => SetActiveView("Dashboard");
    private void BtnMacroEditor_Click(object sender, RoutedEventArgs e) => SetActiveView("MacroEditor");
    private void BtnImageRecognition_Click(object sender, RoutedEventArgs e) => SetActiveView("ImageRecognition");
    private void BtnOcrEngine_Click(object sender, RoutedEventArgs e) => SetActiveView("OcrEngine");
    private void BtnStealthManager_Click(object sender, RoutedEventArgs e) => SetActiveView("StealthManager");
    private void BtnSettings_Click(object sender, RoutedEventArgs e) => SetActiveView("Settings");
    private void BtnAbout_Click(object sender, RoutedEventArgs e) => SetActiveView("About");

    private void BtnNewMacro_Click(object sender, RoutedEventArgs e)
    {
        _currentScript = new MacroScript();
        _actions.Clear();
        SyncScriptToUi();
        SetActiveView("MacroEditor");
        AppendLog("New macro created.");
    }

    // ═══════════════════════════════════════════════════
    //  SETTINGS — HOTKEY UI
    // ═══════════════════════════════════════════════════

    private static readonly string[] ModifierOptions =
        ["Ctrl", "Alt", "Shift", "Ctrl+Alt", "Ctrl+Shift", "Alt+Shift", "Ctrl+Alt+Shift"];

    private static readonly string[] KeyOptions =
        ["F1","F2","F3","F4","F5","F6","F7","F8","F9","F10","F11","F12",
         "A","B","C","D","E","F","G","H","I","J","K","L","M",
         "N","O","P","Q","R","S","T","U","V","W","X","Y","Z"];

    private void InitSettingsUi()
    {
        CmbToggleAppMod.ItemsSource = ModifierOptions;
        CmbToggleAppKey.ItemsSource = KeyOptions;
        CmbToggleTargetMod.ItemsSource = ModifierOptions;
        CmbToggleTargetKey.ItemsSource = KeyOptions;

        CmbToggleAppMod.SelectedItem = HotkeySettings.ModifierToString(_hotkeySettings.ToggleAppModifier);
        CmbToggleAppKey.SelectedItem = HotkeySettings.KeyToString(_hotkeySettings.ToggleAppKey);
        CmbToggleTargetMod.SelectedItem = HotkeySettings.ModifierToString(_hotkeySettings.ToggleTargetModifier);
        CmbToggleTargetKey.SelectedItem = HotkeySettings.KeyToString(_hotkeySettings.ToggleTargetKey);
    }

    private void BtnSaveHotkeys_Click(object sender, RoutedEventArgs e)
    {
        string? appMod = CmbToggleAppMod.SelectedItem as string;
        string? appKey = CmbToggleAppKey.SelectedItem as string;
        string? tgtMod = CmbToggleTargetMod.SelectedItem as string;
        string? tgtKey = CmbToggleTargetKey.SelectedItem as string;

        if (appMod is null || appKey is null || tgtMod is null || tgtKey is null)
        {
            ShowToast("Please select all hotkey fields.", isError: true);
            return;
        }

        UnregisterHotkeys();

        _hotkeySettings.ToggleAppModifier = HotkeySettings.StringToModifier(appMod);
        _hotkeySettings.ToggleAppKey = HotkeySettings.StringToKey(appKey);
        _hotkeySettings.ToggleTargetModifier = HotkeySettings.StringToModifier(tgtMod);
        _hotkeySettings.ToggleTargetKey = HotkeySettings.StringToKey(tgtKey);
        _hotkeySettings.Save();

        RegisterHotkeys();
        ShowToast($"Hotkeys saved: App={_hotkeySettings.ToggleAppDisplay}, Target={_hotkeySettings.ToggleTargetDisplay}", isError: false);
    }

    // ═══════════════════════════════════════════════════
    //  DASHBOARD — MULTI-TASKING HUB (DataGrid)
    // ═══════════════════════════════════════════════════

    private void BtnRefreshDashboard_Click(object sender, RoutedEventArgs e) => LoadDashboard();

    private void LoadDashboard()
    {
        var keepRows = _dashboardRows.Where(r => r.IsRunning).ToList();
        _dashboardRows.Clear();

        var files = ScriptManager.EnumerateSavedScripts().ToList();
        TxtTotalMacros.Text = files.Count.ToString();

        foreach (string filePath in files)
        {
            var existing = keepRows.FirstOrDefault(r => r.FilePath == filePath);
            if (existing is not null)
            {
                _dashboardRows.Add(existing);
                continue;
            }

            MacroScript? script;
            try { script = ScriptManager.Load(filePath); }
            catch { continue; }
            if (script is null) continue;

            string fileStem = Path.GetFileNameWithoutExtension(filePath);
            if (string.IsNullOrWhiteSpace(script.Name) ||
                string.Equals(script.Name, "Untitled Macro", StringComparison.OrdinalIgnoreCase))
            {
                script.Name = string.IsNullOrWhiteSpace(fileStem) ? script.Name : fileStem;
                try { ScriptManager.Save(script, filePath); }
                catch { /* keep in-memory name only */ }
            }

            _dashboardRows.Add(new DashboardRowVm
            {
                FilePath = filePath,
                Script = script,
                TargetWindow = script.TargetWindowTitle,
                RunCount = script.RepeatCount,
                IntervalMinutes = script.IntervalMinutes,
            });
        }

        DashboardEmptyHint.Visibility = _dashboardRows.Count == 0
            ? Visibility.Visible
            : Visibility.Collapsed;

        UpdateProcessBar();
        TxtRunsToday.Text = _runsToday.ToString();
    }

    private void DashRowWindowCombo_DropDownOpened(object sender, EventArgs e)
    {
        if (sender is ComboBox cmb)
        {
            string current = cmb.Text;
            cmb.ItemsSource = GetWindowTitles();
            cmb.Text = current;
        }
    }

    private async void DashboardStart_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button { DataContext: DashboardRowVm row }) return;

        if (string.IsNullOrWhiteSpace(row.TargetWindow))
        {
            ShowToast("Chọn cửa sổ mục tiêu cho macro này.", isError: true);
            return;
        }

        row.Script.TargetWindowTitle = row.TargetWindow;
        row.Script.RepeatCount = row.RunCount;
        row.Script.IntervalMinutes = row.IntervalMinutes;

        IntPtr targetHwnd = ResolveHwnd(row.TargetWindow);
        if (targetHwnd == IntPtr.Zero)
        {
            ShowToast($"Không tìm thấy cửa sổ: \"{row.TargetWindow}\"", isError: true);
            return;
        }
        row.TargetHwnd = targetHwnd;

        if (row.StealthMode && !_hiddenWindows.ContainsKey(targetHwnd))
        {
            string title = Win32Api.GetWindowTitle(targetHwnd);
            StealthHideWindow(targetHwnd, title);
            AppendLog($"[{row.MacroName}] Stealth ON — ẩn cửa sổ mục tiêu.");
        }

        row.Cts = new CancellationTokenSource();
        row.Engine = new MacroEngine();
        row.IsRunning = true;
        row.Status = "Đang chạy";
        UpdateProcessBar();

        row.Engine.Log += msg => Dispatcher.Invoke(() =>
        {
            AppendLog($"[{row.MacroName}] {msg}");
            if (msg.Contains("Waiting")) row.Status = "Đang chờ";
            else if (msg.Contains("Iteration")) row.Status = "Đang chạy";
        });
        row.Engine.ExecutionFinished += () => Dispatcher.Invoke(() =>
        {
            _runsToday++;
            row.IsRunning = false;
            row.Status = "Sẵn sàng";
            RestoreStealthWindow(row);
            UpdateProcessBar();
            AppendLog($"[{row.MacroName}] Hoàn tất.");
        });
        row.Engine.ExecutionFaulted += ex => Dispatcher.Invoke(() =>
        {
            row.IsRunning = false;
            row.Status = "Lỗi";
            RestoreStealthWindow(row);
            UpdateProcessBar();
            AppendLog($"[{row.MacroName}] Lỗi: {ex.Message}");
        });

        try
        {
            AppendLog($"[{row.MacroName}] Bắt đầu trên \"{row.TargetWindow}\"...");
            await row.Engine.ExecuteScriptAsync(row.Script, targetHwnd, row.Cts.Token);
        }
        catch (OperationCanceledException)
        {
            AppendLog($"[{row.MacroName}] Đã dừng.");
        }
        catch (Exception ex)
        {
            AppendLog($"[{row.MacroName}] Lỗi: {ex.Message}");
        }
        finally
        {
            row.IsRunning = false;
            if (row.Status == "Đang chạy" || row.Status == "Đang chờ")
                row.Status = "Đã dừng";
            RestoreStealthWindow(row);
            UpdateProcessBar();
        }
    }

    private void RestoreStealthWindow(DashboardRowVm row)
    {
        if (row.StealthMode && row.TargetHwnd != IntPtr.Zero)
        {
            StealthShowWindow(row.TargetHwnd);
            AppendLog($"[{row.MacroName}] Stealth OFF — hiện lại cửa sổ mục tiêu.");
        }
    }

    private void DashboardStop_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button { DataContext: DashboardRowVm row }) return;
        row.Cts?.Cancel();
        AppendLog($"[{row.MacroName}] Yêu cầu dừng.");
    }

    private async void DashboardRename_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button { DataContext: DashboardRowVm row }) return;

        if (row.IsRunning)
        {
            ShowToast("Dừng macro trước khi đổi tên.", isError: true);
            return;
        }

        var dlg = new RenameMacroDialog(row.Script.Name) { Owner = this };
        if (dlg.ShowDialog() != true || string.IsNullOrWhiteSpace(dlg.NewName))
            return;

        string newDisplayName = dlg.NewName.Trim();
        string oldPath = row.FilePath;
        string dir = Path.GetDirectoryName(oldPath) ?? ScriptManager.DefaultScriptsFolder;

        string newStem = ScriptManager.SanitizeFileStem(newDisplayName);
        string newPath = Path.GetFullPath(Path.Combine(dir, newStem + ".json"));
        string oldFull = Path.GetFullPath(oldPath);

        try
        {
            if (!string.Equals(newPath, oldFull, StringComparison.OrdinalIgnoreCase))
            {
                if (File.Exists(newPath))
                {
                    MessageBox.Show(
                        $"Đã tồn tại file:\n{newPath}",
                        "Đổi tên macro",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }

                File.Move(oldPath, newPath);
                row.FilePath = newPath;
            }

            row.Script.Name = newDisplayName;
            await ScriptManager.SaveAsync(row.Script, row.FilePath);
            row.NotifyScriptMetadataChanged();
            LoadDashboard();
            ShowToast($"Đã đổi tên thành \"{newDisplayName}\".", isError: false);
        }
        catch (Exception ex)
        {
            ShowToast($"Đổi tên thất bại: {ex.Message}", isError: true);
        }
    }

    private void DashboardDelete_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button { DataContext: DashboardRowVm row }) return;

        if (row.IsRunning)
        {
            ShowToast("Dừng macro trước khi xóa.", isError: true);
            return;
        }

        var result = MessageBox.Show(
            $"Bạn có chắc muốn xóa kịch bản \"{row.MacroName}\"?\nFile: {Path.GetFileName(row.FilePath)}",
            "Xác nhận xóa",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning);

        if (result != MessageBoxResult.Yes) return;

        try
        {
            if (File.Exists(row.FilePath))
                File.Delete(row.FilePath);

            AppendLog($"Đã xóa: {row.FilePath}");
            LoadDashboard();
            ShowToast($"Đã xóa \"{row.MacroName}\".", isError: false);
        }
        catch (Exception ex)
        {
            ShowToast($"Lỗi khi xóa: {ex.Message}", isError: true);
        }
    }

    // ═══════════════════════════════════════════════════
    //  PROCESS BAR & STOP ALL
    // ═══════════════════════════════════════════════════

    private void UpdateProcessBar()
    {
        int active = _dashboardRows.Count(r => r.IsRunning);
        TxtActiveProcesses.Text = $"Active: {active} macro{(active == 1 ? "" : "s")} running  |  Hidden: {_hiddenWindows.Count} window(s)";
        TxtActiveThreads.Text = active.ToString();
        ProcessDot.Fill = active > 0
            ? (Brush)FindResource("AccentYellowBrush")
            : (Brush)FindResource("AccentGreenBrush");
    }

    private void BtnStopAllMacros_Click(object sender, RoutedEventArgs e)
    {
        int count = 0;
        foreach (var row in _dashboardRows.Where(r => r.IsRunning))
        {
            row.Cts?.Cancel();
            count++;
        }
        _cts?.Cancel();
        AppendLog($"STOP ALL: đã dừng {count} macro đang chạy.");
    }

    // ═══════════════════════════════════════════════════
    //  MACRO EDITOR — DRAG & DROP
    // ═══════════════════════════════════════════════════

    private void ActionBlock_MouseDown(object sender, MouseButtonEventArgs e)
    {
        if (sender is Border border && border.Child is StackPanel panel)
        {
            string actionType = panel.Tag?.ToString() ?? "Unknown";
            DragDrop.DoDragDrop(border, actionType, DragDropEffects.Copy);
        }
    }

    private void Canvas_DragOver(object sender, DragEventArgs e)
    {
        e.Effects = e.Data.GetDataPresent(DataFormats.StringFormat) ? DragDropEffects.Copy : DragDropEffects.None;
        e.Handled = true;
    }

    private void Canvas_Drop(object sender, DragEventArgs e)
    {
        if (!e.Data.GetDataPresent(DataFormats.StringFormat)) return;
        string actionType = (string)e.Data.GetData(DataFormats.StringFormat);
        MacroAction? action = CreateActionFromType(actionType);
        if (action is null) return;
        _actions.Add(action);
        RebuildCanvas();
        AppendLog($"Added action: {action.DisplayName}");
    }

    private static MacroAction? CreateActionFromType(string actionType) => actionType switch
    {
        "Click" => new ClickAction(),
        "TypeText" => new TypeAction(),
        "Wait" => new WaitAction(),
        "IfImageFound" => new IfImageAction(),
        "IfTextFound" => new IfTextAction(),
        "WebNavigate" => new WebNavigateAction(),
        "WebClick" => new WebClickAction(),
        "WebType" => new WebTypeAction(),
        _ => null,
    };

    // ═══════════════════════════════════════════════════
    //  MACRO EDITOR — CANVAS RENDERING
    // ═══════════════════════════════════════════════════

    private void RebuildCanvas()
    {
        MacroCanvas.Items.Clear();
        CanvasPlaceholder.Visibility = _actions.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        for (int i = 0; i < _actions.Count; i++)
            MacroCanvas.Items.Add(BuildActionCard(_actions[i], i));
    }

    private UIElement BuildActionCard(MacroAction action, int index)
    {
        var (label, color, detail) = action switch
        {
            ClickAction c => ("Click", "#89B4FA", $"X={c.X}  Y={c.Y}"),
            TypeAction t => ("Type Text", "#A6E3A1", $"\"{Truncate(t.Text, 25)}\""),
            WaitAction w => ("Wait", "#F9E2AF", $"{w.Milliseconds}ms"),
            IfImageAction img => ("IF Image Found", "#FAB387",
                Path.GetFileName(img.ImagePath) + (img.ClickOnFound ? " \U0001F3AF Auto-Click" : "")),
            IfTextAction txt => ("IF Text Found", "#B4BEFE", $"\"{Truncate(txt.Text, 25)}\""),
            WebNavigateAction wn => ("Web: Navigate", "#94E2D5", Truncate(wn.Url, 40)),
            WebClickAction wc => ("Web: Click", "#94E2D5", Truncate(wc.CssSelector, 35)),
            WebTypeAction wt => ("Web: Type", "#94E2D5", $"{Truncate(wt.CssSelector, 20)} ← \"{Truncate(wt.TextToType, 15)}\""),
            _ => (action.DisplayName, "#CDD6F4", ""),
        };

        var card = new Border
        {
            Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#313244")),
            CornerRadius = new CornerRadius(6), Padding = new Thickness(12, 8, 12, 8),
            Margin = new Thickness(0, 3, 0, 3), Tag = index,
        };

        var outer = new DockPanel();

        var btnDel = new Button { Content = "X", FontSize = 11, Foreground = Brushes.White, Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F38BA8")), BorderThickness = new Thickness(0), Padding = new Thickness(6, 2, 6, 2), Cursor = Cursors.Hand, VerticalAlignment = VerticalAlignment.Center, Tag = index };
        btnDel.Click += BtnDeleteAction_Click;
        DockPanel.SetDock(btnDel, Dock.Right);
        outer.Children.Add(btnDel);

        var btnEdit = new Button { Content = "Edit", FontSize = 11, Foreground = Brushes.White, Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#89B4FA")), BorderThickness = new Thickness(0), Padding = new Thickness(8, 2, 8, 2), Cursor = Cursors.Hand, VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 6, 0), Tag = index };
        btnEdit.Click += BtnEditAction_Click;
        DockPanel.SetDock(btnEdit, Dock.Right);
        outer.Children.Add(btnEdit);

        var cs = new StackPanel { Orientation = Orientation.Horizontal };
        cs.Children.Add(new System.Windows.Shapes.Ellipse { Width = 8, Height = 8, Fill = new SolidColorBrush((Color)ColorConverter.ConvertFromString(color)), Margin = new Thickness(0, 0, 10, 0), VerticalAlignment = VerticalAlignment.Center });
        cs.Children.Add(new TextBlock { Text = $"[{index}] {label}", Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#CDD6F4")), FontSize = 13, FontWeight = FontWeights.SemiBold, VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 10, 0) });
        if (!string.IsNullOrEmpty(detail))
            cs.Children.Add(new TextBlock { Text = detail, Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#A6ADC8")), FontSize = 11, VerticalAlignment = VerticalAlignment.Center });

        outer.Children.Add(cs);
        card.Child = outer;
        return card;
    }

    private void BtnDeleteAction_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn && btn.Tag is int idx && idx >= 0 && idx < _actions.Count)
        {
            string name = _actions[idx].DisplayName;
            _actions.RemoveAt(idx);
            RebuildCanvas();
            AppendLog($"Removed action [{idx}]: {name}");
        }
    }

    private void BtnEditAction_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button btn || btn.Tag is not int idx || idx < 0 || idx >= _actions.Count) return;
        var action = _actions[idx];
        var dlg = new ActionEditDialog(action) { Owner = this };
        if (dlg.ShowDialog() == true) { RebuildCanvas(); AppendLog($"Edited action [{idx}]: {action.DisplayName}"); }
    }

    private void BtnClearCanvas_Click(object sender, RoutedEventArgs e) { _actions.Clear(); RebuildCanvas(); AppendLog("Canvas cleared."); }

    // ═══════════════════════════════════════════════════
    //  SAVE / LOAD
    // ═══════════════════════════════════════════════════

    private async void BtnSaveMacro_Click(object sender, RoutedEventArgs e)
    {
        SyncUiToScript();
        string suggest = ScriptManager.SanitizeFileStem(_currentScript.Name) + ".json";
        var dlg = new SaveFileDialog
        {
            Filter = "JSON Macro|*.json",
            FileName = suggest,
            InitialDirectory = ScriptManager.DefaultScriptsFolder,
        };
        if (dlg.ShowDialog() != true) return;
        try
        {
            string nameFromFile = Path.GetFileNameWithoutExtension(dlg.FileName);
            if (!string.IsNullOrWhiteSpace(nameFromFile))
                _currentScript.Name = nameFromFile;
            TxtMacroName.Text = _currentScript.Name;

            await ScriptManager.SaveAsync(_currentScript, dlg.FileName);
            AppendLog($"Saved: {dlg.FileName}");
            ShowToast($"Macro saved to {Path.GetFileName(dlg.FileName)}", isError: false);
        }
        catch (Exception ex) { ShowToast($"Save failed: {ex.Message}", isError: true); }
    }

    private async void BtnLoadMacro_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog { Filter = "JSON Macro|*.json", InitialDirectory = ScriptManager.DefaultScriptsFolder };
        if (dlg.ShowDialog() != true) return;
        try
        {
            var script = await ScriptManager.LoadAsync(dlg.FileName);
            if (script is null) { ShowToast("Failed to parse macro file.", isError: true); return; }
            _currentScript = script;
            string baseName = Path.GetFileNameWithoutExtension(dlg.FileName);
            if (string.IsNullOrWhiteSpace(_currentScript.Name) ||
                string.Equals(_currentScript.Name, "Untitled Macro", StringComparison.OrdinalIgnoreCase))
            {
                if (!string.IsNullOrWhiteSpace(baseName))
                    _currentScript.Name = baseName;
            }

            _actions.Clear();
            foreach (var a in _currentScript.Actions) _actions.Add(a);
            SyncScriptToUi();
            RebuildCanvas();
            AppendLog($"Loaded: {dlg.FileName} ({_actions.Count} actions)");
            ShowToast($"Loaded \"{_currentScript.Name}\" ({_actions.Count} actions)", isError: false);
        }
        catch (Exception ex) { ShowToast($"Load failed: {ex.Message}", isError: true); }
    }

    private void SyncScriptToUi()
    {
        TxtMacroName.Text = _currentScript.Name;
        CmbTargetWindow.Text = _currentScript.TargetWindowTitle;
        TxtRepeatCount.Text = _currentScript.RepeatCount.ToString();
    }

    private void SyncUiToScript()
    {
        _currentScript.Name = TxtMacroName.Text.Trim();
        _currentScript.TargetWindowTitle = CmbTargetWindow.Text.Trim();
        _currentScript.Actions = [.. _actions];
        if (int.TryParse(TxtRepeatCount.Text.Trim(), out int repeat)) _currentScript.RepeatCount = repeat;
    }

    // ═══════════════════════════════════════════════════
    //  WINDOW-LIST COMBO HELPERS
    // ═══════════════════════════════════════════════════

    /// <summary>
    /// Resolves a target HWND by partial title, checking the stealth tracker
    /// first so hidden windows are still found.
    /// </summary>
    private IntPtr ResolveHwnd(string partialTitle)
    {
        var hidden = _hiddenWindows
            .FirstOrDefault(kv => kv.Value.Contains(partialTitle, StringComparison.OrdinalIgnoreCase));
        if (hidden.Key != IntPtr.Zero && Win32Api.IsWindow(hidden.Key))
            return hidden.Key;

        return Win32Api.FindWindowByPartialTitle(partialTitle);
    }

    private List<string> GetWindowTitles() =>
        Win32Api.GetAllVisibleWindows()
            .Where(w => w.Title != Title)
            .Select(w => w.Title).ToList();

    private void PopulateCombo(ComboBox cmb)
    {
        string current = cmb.Text;
        cmb.ItemsSource = GetWindowTitles();
        cmb.Text = current;
    }

    private void BtnRefreshWindows_Click(object sender, RoutedEventArgs e) => PopulateCombo(CmbTargetWindow);
    private void CmbTargetWindow_DropDownOpened(object sender, EventArgs e) => PopulateCombo(CmbTargetWindow);

    private void BtnRefreshVisionWindows_Click(object sender, RoutedEventArgs e) => PopulateCombo(CmbVisionWindowTitle);
    private void CmbVisionWindowTitle_DropDownOpened(object sender, EventArgs e) => PopulateCombo(CmbVisionWindowTitle);

    private void BtnRefreshOcrWindows_Click(object sender, RoutedEventArgs e) => PopulateCombo(CmbOcrWindowTitle);
    private void CmbOcrWindowTitle_DropDownOpened(object sender, EventArgs e) => PopulateCombo(CmbOcrWindowTitle);

    // ═══════════════════════════════════════════════════
    //  RUN / STOP MACRO (sidebar buttons — editor macro)
    // ═══════════════════════════════════════════════════

    private async void BtnRunMacro_Click(object sender, RoutedEventArgs e)
    {
        SyncUiToScript();
        if (_actions.Count == 0) { ShowToast("No actions to run.", isError: true); return; }
        if (string.IsNullOrWhiteSpace(_currentScript.TargetWindowTitle)) { ShowToast("Set a Target Window Title.", isError: true); return; }

        IntPtr editorHwnd = ResolveHwnd(_currentScript.TargetWindowTitle);
        if (editorHwnd == IntPtr.Zero) { ShowToast("Target window not found (even in hidden list).", isError: true); return; }

        SetRunningState(true);
        _cts = new CancellationTokenSource();
        _macroEngine = new MacroEngine();
        _macroEngine.Log += msg => Dispatcher.Invoke(() => AppendLog(msg));
        _macroEngine.ActionStarted += (action, idx) => Dispatcher.Invoke(() => TxtStatus.Text = $"Running [{idx}] {action.DisplayName}");
        _macroEngine.ExecutionFinished += () => Dispatcher.Invoke(() => { _runsToday++; SetRunningState(false); ShowToast("Macro completed.", isError: false); UpdateProcessBar(); });
        _macroEngine.ExecutionFaulted += ex => Dispatcher.Invoke(() => { SetRunningState(false); ShowToast($"Error: {ex.Message}", isError: true); UpdateProcessBar(); });

        try
        {
            AppendLog($"Starting macro \"{_currentScript.Name}\"...");
            await _macroEngine.ExecuteScriptAsync(_currentScript, editorHwnd, _cts.Token);
        }
        catch (OperationCanceledException) { ShowToast("Macro stopped by user.", isError: false); }
        catch (Exception ex) { ShowToast($"Error: {ex.Message}", isError: true); }
        finally { SetRunningState(false); UpdateProcessBar(); }
    }

    private void BtnStopMacro_Click(object sender, RoutedEventArgs e) { _cts?.Cancel(); AppendLog("Stop requested."); }

    private void SetRunningState(bool running)
    {
        BtnRunMacro.IsEnabled = !running;
        BtnStopMacro.IsEnabled = running;
        StatusIndicator.Color = running ? (Color)FindResource("AccentYellowColor") : (Color)FindResource("AccentGreenColor");
        TxtStatus.Text = running ? "Running..." : "Ready";
    }

    // ═══════════════════════════════════════════════════
    //  MACRO RECORDING
    // ═══════════════════════════════════════════════════

    private void BtnRecordMacro_Click(object sender, RoutedEventArgs e)
    {
        SyncUiToScript();
        string targetTitle = _currentScript.TargetWindowTitle;
        if (string.IsNullOrWhiteSpace(targetTitle)) { ShowToast("Set a Target Window Title before recording.", isError: true); SetActiveView("MacroEditor"); return; }
        IntPtr hwnd = Win32Api.FindWindowByPartialTitle(targetTitle);
        if (hwnd == IntPtr.Zero) { ShowToast($"Window not found: \"{targetTitle}\".", isError: true); return; }

        _recorder?.Dispose();
        _recorder = new MacroRecorder();
        _recorder.Log += msg => Dispatcher.Invoke(() => AppendLog(msg));

        try { _recorder.StartRecording(hwnd); }
        catch (Exception ex) { ShowToast($"Recording failed: {ex.Message}", isError: true); _recorder.Dispose(); _recorder = null; return; }

        var toolbar = new RecordToolbar(_recorder) { Owner = null };
        toolbar.RecordingFinished += OnRecordingFinished;
        WindowState = WindowState.Minimized;
        toolbar.Show();
    }

    private void OnRecordingFinished(List<MacroAction> recorded)
    {
        WindowState = WindowState.Normal; Activate();
        if (recorded.Count == 0) { ShowToast("No actions recorded.", isError: false); return; }
        foreach (var a in recorded) _actions.Add(a);
        RebuildCanvas();
        SetActiveView("MacroEditor");
        ShowToast($"Recorded {recorded.Count} actions.", isError: false);
    }

    // ═══════════════════════════════════════════════════
    //  IMAGE RECOGNITION TEST
    // ═══════════════════════════════════════════════════

    private void BtnBrowseTemplate_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog { Filter = "Image files|*.png;*.jpg;*.jpeg;*.bmp|All|*.*" };
        if (dlg.ShowDialog() == true) TxtTemplatePath.Text = dlg.FileName;
    }

    private void BtnSnipArea_Click(object sender, RoutedEventArgs e)
    {
        var snip = new SnippingToolWindow();
        if (snip.ShowDialog() == true && !string.IsNullOrEmpty(snip.CapturedFilePath))
        {
            TxtTemplatePath.Text = snip.CapturedFilePath;
            AppendLog($"[Snip] Saved template: {snip.CapturedFilePath}");
        }
    }

    private async void BtnTestVision_Click(object sender, RoutedEventArgs e)
    {
        string windowTitle = CmbVisionWindowTitle.Text.Trim();
        string templatePath = TxtTemplatePath.Text.Trim();
        if (string.IsNullOrEmpty(windowTitle) || string.IsNullOrEmpty(templatePath))
        {
            TxtVisionResult.Text = "Provide both a Window Title and Template Image Path.";
            TxtVisionResult.Foreground = (Brush)FindResource("AccentRedBrush"); return;
        }
        if (!double.TryParse(TxtThreshold.Text.Trim(), out double threshold)) threshold = 0.8;

        TxtVisionResult.Text = "Searching...";
        TxtVisionResult.Foreground = (Brush)FindResource("SubtextBrush");

        try
        {
            var result = await Task.Run(() =>
            {
                IntPtr hwnd = Win32Api.FindWindowByPartialTitle(windowTitle);
                if (hwnd == IntPtr.Zero) return ("Window not found.", false);
                var detailed = VisionEngine.FindImageOnWindowDetailed(hwnd, templatePath);
                if (detailed is null) return ("Template matching returned no data.", false);
                var (loc, conf) = detailed.Value;
                bool found = conf >= threshold;
                return (found ? $"FOUND at ({loc.X}, {loc.Y}) — Confidence: {conf:P1}" : $"NOT FOUND — Best confidence: {conf:P1} (threshold: {threshold:P1})", found);
            });
            TxtVisionResult.Text = result.Item1;
            TxtVisionResult.Foreground = result.Item2 ? (Brush)FindResource("AccentGreenBrush") : (Brush)FindResource("AccentRedBrush");
            AppendLog($"[Vision Test] {result.Item1}");
        }
        catch (Exception ex) { TxtVisionResult.Text = $"Error: {ex.Message}"; TxtVisionResult.Foreground = (Brush)FindResource("AccentRedBrush"); }
    }

    // ═══════════════════════════════════════════════════
    //  OCR TEST
    // ═══════════════════════════════════════════════════

    private async void BtnTestOcr_Click(object sender, RoutedEventArgs e)
    {
        string windowTitle = CmbOcrWindowTitle.Text.Trim();
        if (string.IsNullOrEmpty(windowTitle)) { TxtOcrResult.Text = "Provide a Window Title."; TxtOcrResult.Foreground = (Brush)FindResource("AccentRedBrush"); return; }

        TxtOcrResult.Text = "Extracting text...";
        TxtOcrResult.Foreground = (Brush)FindResource("SubtextBrush");

        try
        {
            string text = await Task.Run(() =>
            {
                IntPtr hwnd = Win32Api.FindWindowByPartialTitle(windowTitle);
                if (hwnd == IntPtr.Zero) throw new InvalidOperationException($"Window not found: \"{windowTitle}\"");
                return VisionEngine.ExtractTextFromWindow(hwnd);
            });
            TxtOcrResult.Text = string.IsNullOrWhiteSpace(text) ? "(no text detected)" : text;
            TxtOcrResult.Foreground = (Brush)FindResource("TextBrush");
            AppendLog($"[OCR Test] Extracted {text.Length} chars.");
        }
        catch (Exception ex) { TxtOcrResult.Text = $"Error: {ex.Message}"; TxtOcrResult.Foreground = (Brush)FindResource("AccentRedBrush"); }
    }

    // ═══════════════════════════════════════════════════
    //  LOG CONSOLE
    // ═══════════════════════════════════════════════════

    private void AppendLog(string message)
    {
        string ts = DateTime.Now.ToString("HH:mm:ss");
        TxtLogConsole.Text += $"[{ts}] {message}\n";
        LogScrollViewer.ScrollToEnd();
    }

    private void BtnClearLog_Click(object sender, RoutedEventArgs e) => TxtLogConsole.Text = string.Empty;

    private void Hyperlink_RequestNavigate(object sender, System.Windows.Navigation.RequestNavigateEventArgs e)
    {
        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(e.Uri.AbsoluteUri) { UseShellExecute = true });
        e.Handled = true;
    }

    // ═══════════════════════════════════════════════════
    //  TOAST
    // ═══════════════════════════════════════════════════

    private async void ShowToast(string message, bool isError)
    {
        AppendLog(isError ? $"[ERROR] {message}" : message);
        TxtStatus.Text = Truncate(message, 60);
        StatusIndicator.Color = isError ? (Color)FindResource("AccentRedColor") : (Color)FindResource("AccentGreenColor");
        await Task.Delay(3000);
        if (_cts is null or { IsCancellationRequested: true })
        {
            StatusIndicator.Color = (Color)FindResource("AccentGreenColor");
            TxtStatus.Text = "Ready";
        }
    }

    // ═══════════════════════════════════════════════════
    //  CLEANUP
    // ═══════════════════════════════════════════════════

    private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
    {
        UnregisterHotkeys();
        ShowAllHiddenWindows();

        foreach (var row in _dashboardRows.Where(r => r.IsRunning)) row.Cts?.Cancel();
        _cts?.Cancel();
        _cts?.Dispose();
        _recorder?.Dispose();

        _trayIcon?.Dispose();
        _trayIcon = null;

        VisionEngine.Shutdown();
    }

    private static string Truncate(string value, int maxLength) =>
        value.Length <= maxLength ? value : string.Concat(value.AsSpan(0, maxLength), "...");
}
