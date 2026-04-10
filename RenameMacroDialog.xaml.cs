using System.Windows;

namespace SmartMacroAI;

public partial class RenameMacroDialog : Window
{
    public string? NewName { get; private set; }

    public RenameMacroDialog(string currentName)
    {
        InitializeComponent();
        TxtNewName.Text = currentName;
        TxtNewName.SelectAll();
        TxtNewName.Focus();
    }

    private void BtnOk_Click(object sender, RoutedEventArgs e)
    {
        NewName = TxtNewName.Text.Trim();
        if (string.IsNullOrWhiteSpace(NewName))
        {
            MessageBox.Show("Please enter a name.", "Rename Macro",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        DialogResult = true;
    }

    private void BtnCancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
    }
}
