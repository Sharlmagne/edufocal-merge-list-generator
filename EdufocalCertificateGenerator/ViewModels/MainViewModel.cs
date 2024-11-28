using System.Configuration;
using System.IO;
using System.Windows.Input;
using EdufocalCertificateGenerator.Models;
using EdufocalCertificateGenerator.Services;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Win32;
using NISInspectorApp.Core;

namespace EdufocalCertificateGenerator.ViewModels;

public class MainViewModel: ViewModel
{

    private string SECTION_NAME = "EmployeeMap";
    private string _mapFilePath;
    private string _fileName;
    private string _aliasEmail;
    private string _courseName;


    public string FileName
    {
        get => _fileName;
        set
        {
            _fileName = value;
            OnPropertyChanged();
        }
    }

    private string MapFilePath
    {
        get => _mapFilePath;
        set
        {
            _mapFilePath = value;
            OnPropertyChanged();
        }
    }

    public string AliasEmail
    {
        get => _aliasEmail;
        set
        {
            _aliasEmail = value;
            OnPropertyChanged();
        }
    }

    public string CourseName
    {
        get => _courseName;
        set
        {
            _courseName = value;
            OnPropertyChanged();
        }
    }


    private Dictionary<string, EmployeeInfo> _employees = new();
    public ICommand UploadEmployeeMapCommand { get; set; }
    public ICommand GenerateCertificateCommand { get; set; }
    public ICommand RemoveEmployeeMapCommand { get; set; }

    public MainViewModel(Configuration AppConfig)
    {
        UploadEmployeeMapCommand = new RelayCommand(UploadEmployeeMapExecute, CanUploadEmployeeMap);
        GenerateCertificateCommand = new RelayCommand(GenerateCertificateExecute, CanGenerateCertificate);
        RemoveEmployeeMapCommand = new RelayCommand(RemoveEmployeeMapExecute, CanRemoveEmployeeMap);

        SetAppConfig(AppConfig);

        // Read Document and generate list of employees
        if (!string.IsNullOrEmpty(MapFilePath))
        {
            LoadEmployeeMap();
        }
    }

    private void SetAppConfig(Configuration AppConfig)
    {
        if (AppConfig.Sections[SECTION_NAME] is null)
        {
            Console.WriteLine("There is no Employee Map");
            AppConfig.Sections.Add(SECTION_NAME, new EmployeeMapSection());
            AppConfig.Save(ConfigurationSaveMode.Modified);
        }
        else
        {
            var userMap = (EmployeeMapSection)AppConfig.Sections[SECTION_NAME];
            MapFilePath = userMap.MapFilePath;
            FileName = userMap.FileName;
        }
    }

    private void LoadEmployeeMap()
    {
        var employees = new EmployeeMap();
        var document = new DocumentReader(MapFilePath);
        document.GenerateList(employees);
        _employees = employees.Employees;
    }

    private void UploadEmployeeMapExecute(object obj)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls|CSV files (*.csv)|*.csv"
        };

        if (dialog.ShowDialog() == true)
        {

            MapFilePath = dialog.FileName;
            FileName = GetFileName(MapFilePath);

            var AppConfig = App.Services.GetRequiredService<Configuration>();

            var userMap = (EmployeeMapSection)AppConfig.Sections[SECTION_NAME];
            userMap.MapFilePath = MapFilePath;
            userMap.FileName = FileName;
            AppConfig.Save(ConfigurationSaveMode.Modified);

            LoadEmployeeMap();
        }
    }

    private void RemoveEmployeeMapExecute(object obj)
    {
        Console.WriteLine("Remove User Map");
        MapFilePath = "";
        FileName = "";

        var AppConfig = App.Services.GetRequiredService<Configuration>();
        var EmployeeMap = (EmployeeMapSection)AppConfig.Sections[SECTION_NAME];

        EmployeeMap.MapFilePath = MapFilePath;
        EmployeeMap.FileName = FileName;

        AppConfig.Save(ConfigurationSaveMode.Modified);
    }

    private bool CanRemoveEmployeeMap(object obj)
    {
        return !string.IsNullOrEmpty(MapFilePath);
    }
    private bool CanUploadEmployeeMap(object obj)
    {
        return true;
    }

    private void GenerateCertificateExecute(object obj)
    {
        if (!_employees.ContainsKey(AliasEmail))
        {
            Console.WriteLine("User not found");
        } else
        {
            Console.WriteLine($"Generate Certificate for {_employees[AliasEmail].FirstName} {_employees[AliasEmail].LastName} - {CourseName}");
        }
    }

    private bool CanGenerateCertificate(object obj)
    {
        // Check if all fields are filled
        if (string.IsNullOrEmpty(AliasEmail) || string.IsNullOrEmpty(MapFilePath) || string.IsNullOrEmpty(CourseName))
        {
            return false;
        }

        return true;
    }

    private string GetFileName(string filePath)
    {
        return Path.GetFileName(filePath);
    }


}