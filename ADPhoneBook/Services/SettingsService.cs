using System.IO;
using System.Text.Json;
using ADPhoneBook.Models;

namespace ADPhoneBook.Services;

public static class SettingsService
{
    private static readonly string SettingsPath =
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ADPhoneBook",
            "settings.json");

    public static AdSettings Load()
    {
        try
        {
            if (File.Exists(SettingsPath))
            {
                var json = File.ReadAllText(SettingsPath);
                return JsonSerializer.Deserialize<AdSettings>(json) ?? new AdSettings();
            }
        }
        catch { /* ignore */ }
        return new AdSettings();
    }

    public static void Save(AdSettings settings)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath)!);
            // Не сохраняем пароль в открытом виде
            var toSave = new AdSettings
            {
                LdapPath       = settings.LdapPath,
                Username       = settings.Username,
                UseCurrentUser = settings.UseCurrentUser
            };
            File.WriteAllText(SettingsPath, JsonSerializer.Serialize(toSave,
                new JsonSerializerOptions { WriteIndented = true }));
        }
        catch { /* ignore */ }
    }
}
