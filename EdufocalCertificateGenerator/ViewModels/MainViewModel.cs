using System.Configuration;
using System.IO;
using System.Windows;
using System.Windows.Input;
using EdufocalCertificateGenerator.Exceptions;
using EdufocalCertificateGenerator.Models;
using EdufocalCertificateGenerator.Services;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Win32;
using NISInspectorApp.Core;
using FileNotFoundException = System.IO.FileNotFoundException;

namespace EdufocalCertificateGenerator.ViewModels;

public class MainViewModel: ViewModel
{

    private string SECTION_NAME = "EmployeeMap";
    private string appDirectory = AppDomain.CurrentDomain.BaseDirectory;
    private string _mapFilePath;
    private string _fileName;
    private string _aliasEmail;
    private string _courseName;
    private string _dateAwarded;


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

    public string DateAwarded
    {
        get => _dateAwarded;
        set
        {
            _dateAwarded = value;
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
        try
        {
            var employees = new EmployeeMap();
            var document = new DocumentReader(MapFilePath);
            document.GenerateList(employees);
            _employees = employees.Employees;
        }
        catch (InvalidFileException ex)
        {
            MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            ClearEmployeeMapPath();
        }
        catch (FileNotFoundException ex)
        {
            MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);

            // Clear the file path
            ClearEmployeeMapPath();
        }
        catch (WorksheetNotFoundException ex)
        {
            MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            ClearEmployeeMapPath();
        }
        catch (Exception ex)
        {
            MessageBox.Show("An unexpected error occurred: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            ClearEmployeeMapPath();
        }
    }

    private void ClearEmployeeMapPath()
    {
        MapFilePath = "";
        FileName = "";
        var AppConfig = App.Services.GetRequiredService<Configuration>();

        var userMap = (EmployeeMapSection)AppConfig.Sections[SECTION_NAME];
        userMap.MapFilePath = MapFilePath;
        userMap.FileName = FileName;
        AppConfig.Save(ConfigurationSaveMode.Modified);
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
        ClearEmployeeMapPath();
    }

    private bool CanRemoveEmployeeMap(object obj)
    {
        // return !string.IsNullOrEmpty(MapFilePath);
        return true;
    }
    private bool CanUploadEmployeeMap(object obj)
    {
        return true;
    }

    private void GenerateCertificateExecute(object obj)
    {
        // Check each field and display error message if empty
        if (string.IsNullOrEmpty(MapFilePath))
        {
            MessageBox.Show("Please upload an employee map (The excel file with all the aliases and the corresponding company emails).", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        if (string.IsNullOrEmpty(AliasEmail))
        {
            MessageBox.Show("Employee Alias email is required", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        if (string.IsNullOrEmpty(CourseName))
        {
            MessageBox.Show("Course Name is required", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        if (string.IsNullOrEmpty(DateAwarded))
        {
            MessageBox.Show("Date Awarded is required", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }




        if (!_employees.TryGetValue(AliasEmail, out var value))
        {
            MessageBox.Show("User not found: Check and update the excel file", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        } else
        {
            // Format Date
            DateTime date = DateTime.Parse(DateAwarded);
            string formattedDate = date.ToString("MMM. d, yyyy");

            // Template Path
            string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "boj_certificate_template.docx");

            DocumentEditor.EditWordDocument(
               templatePath,
value.FirstName + " " + value.LastName,
                CourseName,
                "",
                formattedDate
            );
        }
    }

    private bool CanGenerateCertificate(object obj)
    {
        // Check if all fields are filled
        return true;
    }

    private string GetFileName(string filePath)
    {
        return Path.GetFileName(filePath);
    }
}