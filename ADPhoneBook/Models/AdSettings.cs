namespace ADPhoneBook.Models;

public class AdSettings
{
    public string LdapPath { get; set; } = string.Empty;   // напр. LDAP://DC=company,DC=local
    public string Username { get; set; } = string.Empty;   // домен\пользователь или UPN
    public string Password { get; set; } = string.Empty;
    public bool   UseCurrentUser { get; set; } = true;
}
