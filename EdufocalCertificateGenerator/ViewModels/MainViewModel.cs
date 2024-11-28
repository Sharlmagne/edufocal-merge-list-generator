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

    private string _mapFilePath;
    private string _fileName;
    private string _userAlias;
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

    public string UserAlias
    {
        get => _userAlias;
        set
        {
            _userAlias = value;
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


    private Dictionary<string, UserInfo> _employees = new();
    public ICommand UploadUserMapCommand { get; set; }
    public ICommand GenerateCertificateCommand { get; set; }
    public ICommand RemoveUserMapCommand { get; set; }

    public MainViewModel(Configuration AppConfig)
    {
        UploadUserMapCommand = new RelayCommand(UploadUserMapExecute, CanUploadUserMap);
        GenerateCertificateCommand = new RelayCommand(GenerateCertificateExecute, CanGenerateCertificate);
        RemoveUserMapCommand = new RelayCommand(RemoveUserMapExecute, CanRemoveUserMap);


        if (AppConfig.Sections["UserMap"] is null)
        {
            AppConfig.Sections.Add("UserMap", new UserMapSection());
            AppConfig.Save(ConfigurationSaveMode.Modified);
        }
        else
        {
            var userMap = (UserMapSection)AppConfig.Sections["UserMap"];
            MapFilePath = userMap.MapFilePath;
            FileName = userMap.FileName;
        }

        // Read Document and generate list of employees
        if (!string.IsNullOrEmpty(MapFilePath))
        {
            var employees = new UserMap();
            var document = new DocumentReader(MapFilePath);
            document.GenerateList(employees);
            _employees = employees.Employees;
        }
    }

    private void UploadUserMapExecute(object obj)
    {
        Console.WriteLine("Upload User Map");
        var dialog = new OpenFileDialog
        {
            Filter = "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls|CSV files (*.csv)|*.csv"
        };

        if (dialog.ShowDialog() == true)
        {

            MapFilePath = dialog.FileName;
            FileName = GetFileName(MapFilePath);

            var AppConfig = App.Services.GetRequiredService<Configuration>();

            var userMap = (UserMapSection)AppConfig.Sections["UserMap"];
            userMap.MapFilePath = MapFilePath;
            userMap.FileName = FileName;
            AppConfig.Save(ConfigurationSaveMode.Modified);
        }
    }

    private void RemoveUserMapExecute(object obj)
    {
        Console.WriteLine("Remove User Map");
        MapFilePath = "";
        FileName = "";

        var AppConfig = App.Services.GetRequiredService<Configuration>();
        var userMap = (UserMapSection)AppConfig.Sections["UserMap"];

        userMap.MapFilePath = MapFilePath;
        userMap.FileName = FileName;

        AppConfig.Save(ConfigurationSaveMode.Modified);
    }

    private bool CanRemoveUserMap(object obj)
    {
        return !string.IsNullOrEmpty(MapFilePath);
    }

    private bool CanUploadUserMap(object obj)
    {
        return true;
    }

    private void GenerateCertificateExecute(object obj)
    {
        if (!_employees.ContainsKey(UserAlias))
        {
            Console.WriteLine("User not found");
        } else
        {
            Console.WriteLine($"Generate Certificate for {_employees[UserAlias].FirstName} {_employees[UserAlias].LastName} - {CourseName}");
        }
    }

    private bool CanGenerateCertificate(object obj)
    {
        // Check if all fields are filled
        if (string.IsNullOrEmpty(UserAlias) || string.IsNullOrEmpty(MapFilePath) || string.IsNullOrEmpty(CourseName))
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