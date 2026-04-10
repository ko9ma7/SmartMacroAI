using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;
using SmartMacroAI.Core;
using SmartMacroAI.Models;

namespace SmartMacroAI;

public partial class RecordToolbar : Window
{
    private readonly MacroRecorder _recorder;
    private readonly DispatcherTimer _uiTimer;
    private bool _dotVisible = true;

    /// <summary>
    /// Fires when recording finishes. Passes the captured action list.
    /// </summary>
    public event Action<List<MacroAction>>? RecordingFinished;

    public RecordToolbar(MacroRecorder recorder)
    {
        InitializeComponent();

        _recorder = recorder;
        _recorder.ActionRecorded += OnActionRecorded;
        _recorder.StopKeyPressed += OnStopKeyPressed;

        _uiTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(500) };
        _uiTimer.Tick += UiTimer_Tick;
        _uiTimer.Start();
    }

    // ═══════════════════════════════════════════════
    //  WINDOW EVENTS
    // ═══════════════════════════════════════════════

    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
        Left = (SystemParameters.PrimaryScreenWidth - ActualWidth) / 2;
        Top = 16;
    }

    private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        DragMove();
    }

    // ═══════════════════════════════════════════════
    //  UI TIMER (blinking dot + elapsed time)
    // ═══════════════════════════════════════════════

    private void UiTimer_Tick(object? sender, EventArgs e)
    {
        if (!_recorder.IsRecording) return;

        TxtTimer.Text = _recorder.Elapsed.ToString(@"mm\:ss");

        _dotVisible = !_dotVisible;
        RecordDot.Opacity = _dotVisible ? 1.0 : 0.15;
    }

    // ═══════════════════════════════════════════════
    //  RECORDER EVENTS
    // ═══════════════════════════════════════════════

    private void OnActionRecorded(int count)
    {
        Dispatcher.Invoke(() =>
        {
            TxtActionCount.Text = $"{count} action{(count == 1 ? "" : "s")}";
        });
    }

    private void OnStopKeyPressed()
    {
        Dispatcher.BeginInvoke(FinishRecording);
    }

    // ═══════════════════════════════════════════════
    //  STOP
    // ═══════════════════════════════════════════════

    private void BtnStop_Click(object sender, RoutedEventArgs e)
    {
        FinishRecording();
    }

    private void FinishRecording()
    {
        _uiTimer.Stop();

        _recorder.ActionRecorded -= OnActionRecorded;
        _recorder.StopKeyPressed -= OnStopKeyPressed;

        var actions = _recorder.StopRecording();
        RecordingFinished?.Invoke(actions);
        Close();
    }
}
