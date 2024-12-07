namespace EdufocalMergeListGenerator.Models;

public class EmployeesList
{
    public Dictionary<string, EmployeeInfo> Employees { get; set; }

    public EmployeesList()
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