using System.Collections.ObjectModel;
using System.Windows;
using SmartMacroAI.Models;

namespace SmartMacroAI.Core;

/// <summary>
/// Application-wide notification service for user notifications.
/// Created by Phạm Duy – Giải pháp tự động hóa thông minh.
/// </summary>
public sealed class NotificationService
{
    private const int MaxNotifications = 50;
    private static NotificationService? _instance;
    private static readonly object Lock = new();

    public static NotificationService Instance
    {
        get
        {
            if (_instance == null)
            {
                lock (Lock)
                {
                    _instance ??= new NotificationService();
                }
            }
            return _instance;
        }
    }

    public ObservableCollection<AppNotification> Notifications { get; } = new();

    public event Action<AppNotification>? NotificationAdded;

    public int UnreadCount => Notifications.Count(n => !n.IsRead);

    public void Push(NotificationType type, string title, string message, string? macroName = null)
    {
        var notification = new AppNotification
        {
            Type = type,
            Title = title,
            Message = message,
            MacroName = macroName
        };

        Application.Current?.Dispatcher.Invoke(() =>
        {
            Notifications.Insert(0, notification);
            if (Notifications.Count > MaxNotifications)
                Notifications.RemoveAt(Notifications.Count - 1);
        });

        NotificationAdded?.Invoke(notification);
    }

    public void PushSuccess(string title, string message, string? macroName = null)
        => Push(NotificationType.Success, title, message, macroName);

    public void PushError(string title, string message, string? macroName = null)
        => Push(NotificationType.Error, title, message, macroName);

    public void PushWarning(string title, string message, string? macroName = null)
        => Push(NotificationType.Warning, title, message, macroName);

    public void PushInfo(string title, string message, string? macroName = null)
        => Push(NotificationType.Info, title, message, macroName);

    public void PushSchedule(string title, string message, string? macroName = null)
        => Push(NotificationType.Schedule, title, message, macroName);

    public void MarkAllRead()
    {
        Application.Current?.Dispatcher.Invoke(() =>
        {
            foreach (var n in Notifications)
                n.IsRead = true;
        });
    }

    public void Clear()
    {
        Application.Current?.Dispatcher.Invoke(() => Notifications.Clear());
    }

    public void MarkRead(string notificationId)
    {
        Application.Current?.Dispatcher.Invoke(() =>
        {
            var notification = Notifications.FirstOrDefault(n => n.Id == notificationId);
            if (notification != null)
                notification.IsRead = true;
        });
    }
}
