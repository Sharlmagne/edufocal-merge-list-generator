namespace EdufocalCertificateGenerator.Models;

public class UserMap
{
    public Dictionary<string, UserInfo> Employees { get; set; }

    public UserMap()
    {
        Employees = new Dictionary<string, UserInfo>();
    }

    public void AddUser(string aliasEmail, string firstName, string lastName, string companyEmail)
    {
        var userInfo = new UserInfo
        {
            FirstName = firstName,
            LastName = lastName,
            CompanyEmail = companyEmail
        };

        Employees[aliasEmail] = userInfo;
    }
}