namespace EdufocalCertificateGenerator.Models;

public class EmployeeMap
{
    public Dictionary<string, EmployeeInfo> Employees { get; set; }

    public EmployeeMap()
    {
        Employees = new Dictionary<string, EmployeeInfo>();
    }

    public void AddUser(string aliasEmail, string firstName, string lastName, string companyEmail)
    {
        var userInfo = new EmployeeInfo
        {
            FirstName = firstName,
            LastName = lastName,
            CompanyEmail = companyEmail
        };

        Employees[aliasEmail] = userInfo;
    }
}