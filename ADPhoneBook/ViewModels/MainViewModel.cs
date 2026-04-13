using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using ADPhoneBook.Helpers;
using ADPhoneBook.Models;
using ADPhoneBook.Services;

namespace ADPhoneBook.ViewModels;

public class MainViewModel : INotifyPropertyChanged
{
    private readonly IAdService _adService = new AdService();

    // ─── Backing fields ───────────────────────────────────────────────────────

    private ObservableCollection<Employee> _employees = new();
    private ICollectionView? _employeesView;
    private string _searchText     = string.Empty;
    private string _selectedDept   = "Все отделы";
    private ObservableCollection<string> _departments = new();
    private bool   _isLoading;
    private string _statusText     = "Готов";
    private Employee? _selectedEmployee;
    private AdSettings _settings   = new();
    private CancellationTokenSource? _cts;

    // ─── Public properties ────────────────────────────────────────────────────

    public ICollectionView? EmployeesView
    {
        get => _employeesView;
        private set { _employeesView = value; OnPropertyChanged(); }
    }

    public string SearchText
    {
        get => _searchText;
        set
        {
            if (_searchText == value) return;
            _searchText = value;
            OnPropertyChanged();
            _employeesView?.Refresh();
            UpdateStatus();
        }
    }

    public string SelectedDepartment
    {
        get => _selectedDept;
        set
        {
            if (_selectedDept == value) return;
            _selectedDept = value;
            OnPropertyChanged();
            _employeesView?.Refresh();
            UpdateStatus();
        }
    }

    public ObservableCollection<string> Departments
    {
        get => _departments;
        private set { _departments = value; OnPropertyChanged(); }
    }

    public bool IsLoading
    {
        get => _isLoading;
        private set { _isLoading = value; OnPropertyChanged(); OnPropertyChanged(nameof(IsNotLoading)); }
    }

    public bool IsNotLoading => !_isLoading;

    public string StatusText
    {
        get => _statusText;
        private set { _statusText = value; OnPropertyChanged(); }
    }

    public Employee? SelectedEmployee
    {
        get => _selectedEmployee;
        set { _selectedEmployee = value; OnPropertyChanged(); }
    }

    public AdSettings Settings
    {
        get => _settings;
        set { _settings = value; OnPropertyChanged(); }
    }

    // ─── Commands ─────────────────────────────────────────────────────────────

    public ICommand LoadCommand        { get; }
    public ICommand CancelCommand      { get; }
    public ICommand ExportExcelCommand { get; }
    public ICommand ExportCsvCommand   { get; }
    public ICommand ClearSearchCommand { get; }
    public ICommand OpenSettingsCommand{ get; }

    // ─── Constructor ──────────────────────────────────────────────────────────

    public MainViewModel()
    {
        _settings = SettingsService.Load();

        LoadCommand         = new RelayCommand(async () => await LoadEmployeesAsync(), () => IsNotLoading);
        CancelCommand       = new RelayCommand(CancelLoad, () => IsLoading);
        ExportExcelCommand  = new RelayCommand(ExportToExcel,  () => _employees.Count > 0 && IsNotLoading);
        ExportCsvCommand    = new RelayCommand(ExportToCsv,    () => _employees.Count > 0 && IsNotLoading);
        ClearSearchCommand  = new RelayCommand(() => SearchText = string.Empty);
        OpenSettingsCommand = new RelayCommand(OpenSettings);
    }

    // ─── Load ─────────────────────────────────────────────────────────────────

    private async Task LoadEmployeesAsync()
    {
        _cts = new CancellationTokenSource();
        IsLoading  = true;
        StatusText = "Загрузка данных из Active Directory…";

        try
        {
            var list = await _adService.GetEmployeesAsync(_settings, _cts.Token);

            _employees = new ObservableCollection<Employee>(list);

            var view = CollectionViewSource.GetDefaultView(_employees);
            view.Filter = ApplyFilter;
            EmployeesView = view;

            RebuildDepartments(list);
            SettingsService.Save(_settings);
            UpdateStatus();
        }
        catch (OperationCanceledException)
        {
            StatusText = "Загрузка отменена.";
        }
        catch (Exception ex)
        {
            StatusText = $"Ошибка: {ex.Message}";
            MessageBox.Show(
                $"Не удалось подключиться к Active Directory:\n\n{ex.Message}\n\nПроверьте настройки подключения.",
                "Ошибка подключения",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            IsLoading = false;
            _cts?.Dispose();
            _cts = null;
        }
    }

    private void CancelLoad() => _cts?.Cancel();

    // ─── Filter ───────────────────────────────────────────────────────────────

    private bool ApplyFilter(object obj)
    {
        if (obj is not Employee emp) return false;

        bool deptMatch = SelectedDepartment == "Все отделы"
            || emp.Department.Equals(SelectedDepartment, StringComparison.OrdinalIgnoreCase);

        if (!deptMatch) return false;

        if (string.IsNullOrWhiteSpace(SearchText)) return true;

        var q = SearchText.Trim();
        return emp.DisplayName.Contains(q, StringComparison.OrdinalIgnoreCase)
            || emp.Department.Contains(q, StringComparison.OrdinalIgnoreCase)
            || emp.Title.Contains(q, StringComparison.OrdinalIgnoreCase)
            || emp.Email.Contains(q, StringComparison.OrdinalIgnoreCase)
            || emp.WorkPhone.Contains(q, StringComparison.OrdinalIgnoreCase)
            || emp.MobilePhone.Contains(q, StringComparison.OrdinalIgnoreCase);
    }

    private void RebuildDepartments(List<Employee> list)
    {
        var depts = list
            .Select(e => e.Department)
            .Where(d => !string.IsNullOrWhiteSpace(d))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(d => d)
            .ToList();

        var col = new ObservableCollection<string> { "Все отделы" };
        foreach (var d in depts) col.Add(d);
        Departments = col;
        SelectedDepartment = "Все отделы";
    }

    private void UpdateStatus()
    {
        if (_employeesView is null) { StatusText = "Готов"; return; }
        int visible = _employeesView.Cast<object>().Count();
        StatusText = $"Показано: {visible} из {_employees.Count} сотрудников";
    }

    // ─── Export ───────────────────────────────────────────────────────────────

    private void ExportToExcel()
    {
        var dlg = new Microsoft.Win32.SaveFileDialog
        {
            Title      = "Экспорт в Excel",
            Filter     = "Excel файл (*.xlsx)|*.xlsx",
            FileName   = "phonebook.xlsx",
            DefaultExt = ".xlsx"
        };
        if (dlg.ShowDialog() != true) return;

        try
        {
            var items = GetFilteredEmployees();
            ExportService.ExportToExcel(items, dlg.FileName);
            StatusText = $"Экспортировано {items.Count} записей → {dlg.FileName}";
            if (MessageBox.Show("Файл сохранён. Открыть?", "Экспорт", MessageBoxButton.YesNo,
                    MessageBoxImage.Question) == MessageBoxResult.Yes)
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(dlg.FileName) { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Ошибка экспорта:\n{ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void ExportToCsv()
    {
        var dlg = new Microsoft.Win32.SaveFileDialog
        {
            Title      = "Экспорт в CSV",
            Filter     = "CSV файл (*.csv)|*.csv",
            FileName   = "phonebook.csv",
            DefaultExt = ".csv"
        };
        if (dlg.ShowDialog() != true) return;

        try
        {
            var items = GetFilteredEmployees();
            ExportService.ExportToCsv(items, dlg.FileName);
            StatusText = $"Экспортировано {items.Count} записей → {dlg.FileName}";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Ошибка экспорта:\n{ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private List<Employee> GetFilteredEmployees()
    {
        return _employeesView?.Cast<Employee>().ToList() ?? new List<Employee>();
    }

    // ─── Settings ─────────────────────────────────────────────────────────────

    private void OpenSettings()
    {
        var win = new Views.SettingsWindow(_settings) { Owner = Application.Current.MainWindow };
        if (win.ShowDialog() == true)
        {
            Settings = win.Result;
        }
    }

    // ─── INotifyPropertyChanged ───────────────────────────────────────────────

    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}
