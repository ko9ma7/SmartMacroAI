// Created by Phạm Duy – Giải pháp tự động hóa thông minh.

using System.Windows;
using System.Windows.Input;
using SmartMacroAI.Core;
using SmartMacroAI.Models;

namespace SmartMacroAI;

public partial class MacroLockDialog : Window
{
    public string? NewPasswordHash { get; private set; }
    public bool RemoveLock { get; private set; }

    public MacroLockDialog(MacroScript script)
    {
        InitializeComponent();
        if (!string.IsNullOrEmpty(script.PasswordHash))
        {
            BtnRemoveLock.Visibility = Visibility.Visible;
        }
    }

    private void BtnSetLock_Click(object sender, RoutedEventArgs e)
    {
        string pwd = PwdNew.Password;
        string confirm = PwdConfirm.Password;

        if (string.IsNullOrWhiteSpace(pwd))
        {
            TxtError.Text = "Mật khẩu không được trống.";
            return;
        }

        if (pwd.Length < 4)
        {
            TxtError.Text = "Mật khẩu phải có ít nhất 4 ký tự.";
            return;
        }

        if (pwd != confirm)
        {
            TxtError.Text = "Mật khẩu xác nhận không khớp.";
            return;
        }

        NewPasswordHash = MacroLockService.HashPassword(pwd);
        RemoveLock = false;
        DialogResult = true;
    }

    private void BtnRemoveLock_Click(object sender, RoutedEventArgs e)
    {
        RemoveLock = true;
        DialogResult = true;
    }

    private void BtnCancel_Click(object sender, RoutedEventArgs e)
        => DialogResult = false;
}
