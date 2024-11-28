using System.Configuration;

namespace EdufocalCertificateGenerator.Models;

public class EmployeeMapSection : ConfigurationSection
{
    [ConfigurationProperty("MapFilePath", DefaultValue = "")]
    public string MapFilePath
    {
        get => (string)this["MapFilePath"];
        set => this["MapFilePath"] = value;
    }

    [ConfigurationProperty("FileName", DefaultValue = "")]
    public string FileName
    {
        get => (string)this["FileName"];
        set => this["FileName"] = value;
    }
}