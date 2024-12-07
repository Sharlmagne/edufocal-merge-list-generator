using System.Configuration;

namespace EdufocalMergeListGenerator.Models;

public class EmployeesListSection : ConfigurationSection
{
    private const string _name = "EmployeesListFilePath";
    private const string _fileName = "EmployeesListFileName";

    [ConfigurationProperty(_name, DefaultValue = "")]
    public string MapFilePath
    {
        get => (string)this[_name];
        set => this[_name] = value;
    }

    [ConfigurationProperty(_fileName, DefaultValue = "")]
    public string FileName
    {
        get => (string)this[_fileName];
        set => this[_fileName] = value;
    }
}