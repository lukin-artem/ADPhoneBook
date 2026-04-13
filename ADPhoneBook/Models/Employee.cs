namespace ADPhoneBook.Models;

public class Employee
{
    public string DisplayName   { get; set; } = string.Empty;
    public string Department    { get; set; } = string.Empty;
    public string Title         { get; set; } = string.Empty;
    public string WorkPhone     { get; set; } = string.Empty;
    public string MobilePhone   { get; set; } = string.Empty;
    public string Email         { get; set; } = string.Empty;

    /// <summary>Инициалы для аватара</summary>
    public string Initials
    {
        get
        {
            if (string.IsNullOrWhiteSpace(DisplayName)) return "?";
            var parts = DisplayName.Trim().Split(' ', StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 1) return parts[0][..1].ToUpper();
            return (parts[0][..1] + parts[1][..1]).ToUpper();
        }
    }
}
