using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using SmartMacroAI.Core;
using SmartMacroAI.Models;

namespace SmartMacroAI;

/// <summary>
/// Run history dialog showing macro execution history.
/// Created by Phạm Duy – Giải pháp tự động hóa thông minh.
/// </summary>
public partial class RunHistoryDialog : Window
{
    private readonly RunHistoryService _historyService = new();
    private readonly string? _macroNameFilter;
    private readonly ObservableCollection<MacroRunRecord> _records = [];

    public RunHistoryDialog(string? macroName = null)
    {
        InitializeComponent();
        _macroNameFilter = macroName;
        HistoryGrid.ItemsSource = _records;

        if (!string.IsNullOrWhiteSpace(macroName))
        {
            Title = $"📊 Lịch sử: {macroName}";
            TxtMacroNameFilter.Text = $"Macro: {macroName}";
        }
        else
        {
            Title = "📊 Toàn bộ lịch sử chạy";
            TxtMacroNameFilter.Text = "Hiển thị tất cả macro";
        }

        LoadHistory();
    }

    private void LoadHistory()
    {
        _records.Clear();

        List<MacroRunRecord> records;
        if (string.IsNullOrWhiteSpace(_macroNameFilter))
            records = _historyService.LoadAll();
        else
            records = _historyService.Load(_macroNameFilter);

        foreach (var record in records)
            _records.Add(record);
    }

    private void HistoryGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (HistoryGrid.SelectedItem is MacroRunRecord record)
        {
            TxtLogPreview.Text = string.IsNullOrWhiteSpace(record.LogSnapshot)
                ? "(Không có log)"
                : record.LogSnapshot;
        }
        else
        {
            TxtLogPreview.Text = "(Chọn một bản ghi để xem log)";
        }
    }

    private void BtnViewLog_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn && btn.Tag is MacroRunRecord record)
        {
            if (string.IsNullOrWhiteSpace(record.LogSnapshot))
            {
                MessageBox.Show("Không có log cho bản ghi này.", "Log", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // Show log in a simple text window or copy to clipboard
            var logWindow = new Window
            {
                Title = $"Log: {record.MacroName} ({record.StartTime:yyyy-MM-dd HH:mm:ss})",
                Width = 600,
                Height = 400,
                Background = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#1E1E2E")),
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = this
            };

            var grid = new Grid();
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

            var scrollViewer = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                Margin = new Thickness(12)
            };

            var textBlock = new TextBlock
            {
                Text = record.LogSnapshot,
                Foreground = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#CDD6F4")),
                FontFamily = new System.Windows.Media.FontFamily("Consolas"),
                FontSize = 11,
                TextWrapping = TextWrapping.NoWrap
            };

            scrollViewer.Content = textBlock;
            Grid.SetRow(scrollViewer, 1);

            grid.Children.Add(scrollViewer);
            logWindow.Content = grid;
            logWindow.ShowDialog();
        }
    }

    private void BtnViewScreenshot_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn && btn.Tag is MacroRunRecord record)
        {
            if (string.IsNullOrWhiteSpace(record.ScreenshotPath) || !File.Exists(record.ScreenshotPath))
            {
                MessageBox.Show("Ảnh chụp màn hình không tồn tại.", "Screenshot", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = record.ScreenshotPath,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Không thể mở ảnh: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    private void BtnClearHistory_Click(object sender, RoutedEventArgs e)
    {
        string message = string.IsNullOrWhiteSpace(_macroNameFilter)
            ? "Bạn có chắc muốn xóa toàn bộ lịch sử chạy?"
            : $"Bạn có chắc muốn xóa lịch sử của macro '{_macroNameFilter}'?";

        var result = MessageBox.Show(message, "Xác nhận xóa", MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (result == MessageBoxResult.Yes)
        {
            _historyService.Clear(_macroNameFilter);
            LoadHistory();
            TxtLogPreview.Text = "(Đã xóa lịch sử)";
        }
    }
}
