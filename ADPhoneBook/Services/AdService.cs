using System.DirectoryServices;
using ADPhoneBook.Models;

namespace ADPhoneBook.Services;

public interface IAdService
{
    Task<List<Employee>> GetEmployeesAsync(AdSettings settings, CancellationToken ct = default);
}

public class AdService : IAdService
{
    private static readonly string[] LdapProps =
    {
        "displayName", "telephoneNumber", "mobile",
        "mail", "department", "title"
    };

    public Task<List<Employee>> GetEmployeesAsync(AdSettings settings, CancellationToken ct = default)
    {
        return Task.Run(() => QueryAd(settings, ct), ct);
    }

    private List<Employee> QueryAd(AdSettings settings, CancellationToken ct)
    {
        var result = new List<Employee>();

        DirectoryEntry entry;

        if (settings.UseCurrentUser || string.IsNullOrWhiteSpace(settings.Username))
        {
            entry = string.IsNullOrWhiteSpace(settings.LdapPath)
                ? new DirectoryEntry()
                : new DirectoryEntry(settings.LdapPath);
        }
        else
        {
            entry = string.IsNullOrWhiteSpace(settings.LdapPath)
                ? new DirectoryEntry("", settings.Username, settings.Password,
                      AuthenticationTypes.Secure)
                : new DirectoryEntry(settings.LdapPath, settings.Username, settings.Password,
                      AuthenticationTypes.Secure);
        }

        using (entry)
        using (var searcher = new DirectorySearcher(entry))
        {
            searcher.Filter      = "(&(objectClass=user)(objectCategory=person)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))";
            searcher.PageSize    = 1000;
            searcher.SizeLimit   = 0;
            searcher.SearchScope = SearchScope.Subtree;

            foreach (var prop in LdapProps)
                searcher.PropertiesToLoad.Add(prop);

            using var results = searcher.FindAll();

            foreach (SearchResult sr in results)
            {
                ct.ThrowIfCancellationRequested();

                result.Add(new Employee
                {
                    DisplayName  = GetProp(sr, "displayName"),
                    WorkPhone    = GetProp(sr, "telephoneNumber"),
                    MobilePhone  = GetProp(sr, "mobile"),
                    Email        = GetProp(sr, "mail"),
                    Department   = GetProp(sr, "department"),
                    Title        = GetProp(sr, "title"),
                });
            }
        }

        return result
            .Where(e => !string.IsNullOrWhiteSpace(e.DisplayName)
            && !string.IsNullOrWhiteSpace(e.WorkPhone))
            .OrderBy(e => e.DisplayName)
            .ToList();
    }

    private static string GetProp(SearchResult sr, string prop)
    {
        if (sr.Properties.Contains(prop) && sr.Properties[prop].Count > 0)
            return sr.Properties[prop][0]?.ToString()?.Trim() ?? string.Empty;
        return string.Empty;
    }
}
