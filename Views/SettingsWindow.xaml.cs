using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using ADPhoneBook.Models;

namespace ADPhoneBook.Views;

public partial class SettingsWindow : Window, INotifyPropertyChanged
{
    // ── ViewModel-like state ──────────────────────────────────────────────────
    private string _ldapPath     = string.Empty;
    private string _username     = string.Empty;
    private bool   _useCurrentUser = true;

    public string LdapPath
    {
        get => _ldapPath;
        set { _ldapPath = value; OnPropertyChanged(); }
    }
    public string Username
    {
        get => _username;
        set { _username = value; OnPropertyChanged(); }
    }
    public bool UseCurrentUser
    {
        get => _useCurrentUser;
        set { _useCurrentUser = value; OnPropertyChanged(); OnPropertyChanged(nameof(IsNotCurrentUser)); }
    }
    public bool IsNotCurrentUser => !_useCurrentUser;

    public AdSettings Result { get; private set; } = new();

    // ── Constructor ───────────────────────────────────────────────────────────
    public SettingsWindow(AdSettings current)
    {
        InitializeComponent();
        DataContext = this;

        LdapPath       = current.LdapPath;
        Username       = current.Username;
        UseCurrentUser = current.UseCurrentUser;
    }

    // ── Buttons ───────────────────────────────────────────────────────────────
    private void Save_Click(object sender, RoutedEventArgs e)
    {
        Result = new AdSettings
        {
            LdapPath       = LdapPath.Trim(),
            Username       = Username.Trim(),
            Password       = PwdBox.Password,
            UseCurrentUser = UseCurrentUser
        };
        DialogResult = true;
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
    }

    // ── INotifyPropertyChanged ────────────────────────────────────────────────
    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}
