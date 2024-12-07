using System.Configuration;
using System.IO;
using System.Windows;
using System.Windows.Input;
using Bytescout.Spreadsheet;
using EdufocalMergeListGenerator.Exceptions;
using EdufocalMergeListGenerator.Models;
using EdufocalMergeListGenerator.Services;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Win32;
using NISInspectorApp.Core;
using FileNotFoundException = System.IO.FileNotFoundException;

namespace EdufocalMergeListGenerator.ViewModels;

public class MainViewModel: ViewModel
{

    private string SECTION_NAME = "EmployeesList";
    private string appDirectory = AppDomain.CurrentDomain.BaseDirectory;
    private string _employeesListFilePath;
    private string _employeesListFileName;
    private string _aliasListFilePath;
    private string _aliasListFileName;

    public string EmployeesListFileName
    {
        get => _employeesListFileName;
        set
        {
            _employeesListFileName = value;
            OnPropertyChanged();
        }
    }

    private string EmployeesListFilePath
    {
        get => _employeesListFilePath;
        set
        {
            _employeesListFilePath = value;
            OnPropertyChanged();
        }
    }

    public string AliasListFileName
    {
        get => _aliasListFileName;
        set
        {
            _aliasListFileName = value;
            OnPropertyChanged();
        }
    }

    private string AliasListFilePath
    {
        get => _aliasListFilePath;
        set
        {
            _aliasListFilePath = value;
            OnPropertyChanged();
        }
    }

    private Dictionary<string, EmployeeInfo> Employees { get; set; } = new();
    private Worksheet AliasListWorksheet { get; set; }

    public ICommand UploadEmployeesListCommand { get; set; }
    public ICommand RemoveEmployeesListCommand { get; set; }
    public ICommand UploadAliasListCommand { get; set; }
    public ICommand RemoveAliasListCommand { get; set; }
    public ICommand GenerateMergeListCommand { get; set; }


    public MainViewModel(Configuration AppConfig)
    {
        UploadEmployeesListCommand = new RelayCommand(UploadEmployeesListExecute, _ => true);
        RemoveEmployeesListCommand = new RelayCommand(RemoveEmployeesListExecute, _ => true);
        UploadAliasListCommand = new RelayCommand(UploadAliasListExecute, _ => true);
        RemoveAliasListCommand = new RelayCommand(RemoveAliasListExecute, _ => true);
        GenerateMergeListCommand = new RelayCommand(GenerateMergeListExecute, _ => true);

        SetAppConfig(AppConfig);

        // Read Document and generate list of employees
        if (!string.IsNullOrEmpty(EmployeesListFilePath))
        {
            LoadEmployeeMap();
        }
    }

    private void SetAppConfig(Configuration AppConfig)
    {
        if (AppConfig.Sections[SECTION_NAME] is null)
        {
            Console.WriteLine("There is no Employee Map");
            AppConfig.Sections.Add(SECTION_NAME, new EmployeesListSection());
            AppConfig.Save(ConfigurationSaveMode.Modified);
        }
        else
        {
            var userMap = (EmployeesListSection)AppConfig.Sections[SECTION_NAME];
            EmployeesListFilePath = userMap.MapFilePath;
            EmployeesListFileName = userMap.FileName;
        }
    }

    private void LoadEmployeeMap()
    {
        try
        {
            var employees = new EmployeesList();
            var document = new DocumentReader(EmployeesListFilePath);
            document.GenerateEmployeesList(employees);
            Employees = employees.Employees;
        }
        catch (InvalidFileException ex)
        {
            MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            ClearEmployeesListPath();
        }
        catch (FileNotFoundException ex)
        {
            MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);

            // Clear the file path
            ClearEmployeesListPath();
        }
        catch (WorksheetNotFoundException ex)
        {
            MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            ClearEmployeesListPath();
        }
        catch (Exception ex)
        {
            MessageBox.Show("An unexpected error occurred: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            ClearEmployeesListPath();
        }
    }

    private void ClearEmployeesListPath()
    {
        EmployeesListFilePath = "";
        EmployeesListFileName = "";
        var AppConfig = App.Services.GetRequiredService<Configuration>();

        var userMap = (EmployeesListSection)AppConfig.Sections[SECTION_NAME];
        userMap.MapFilePath = EmployeesListFilePath;
        userMap.FileName = EmployeesListFileName;
        AppConfig.Save(ConfigurationSaveMode.Modified);
    }

    private void ClearAliasListPath()
    {
        AliasListFilePath = "";
        AliasListFileName = "";
    }

    private void UploadEmployeesListExecute(object obj)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls|CSV files (*.csv)|*.csv"
        };

        if (dialog.ShowDialog() == true)
        {

            EmployeesListFilePath = dialog.FileName;
            EmployeesListFileName = GetFileName(EmployeesListFilePath);

            var AppConfig = App.Services.GetRequiredService<Configuration>();

            var userMap = (EmployeesListSection)AppConfig.Sections[SECTION_NAME];
            userMap.MapFilePath = EmployeesListFilePath;
            userMap.FileName = EmployeesListFileName;
            AppConfig.Save(ConfigurationSaveMode.Modified);

            LoadEmployeeMap();
        }
    }

    private void RemoveEmployeesListExecute(object obj)
    {
        Console.WriteLine("Remove User Map");
        ClearEmployeesListPath();
    }

    private void UploadAliasListExecute(object obj)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls|CSV files (*.csv)|*.csv"
        };

        if (dialog.ShowDialog() == true)
        {

            AliasListFilePath = dialog.FileName;
            AliasListFileName = GetFileName(AliasListFilePath);

            try
            {
                var Document = new DocumentReader(AliasListFilePath);
                Document.ValidateAliasList();
                AliasListWorksheet = Document.Document.Workbook.Worksheets.ByName("Sheet1");
            }
            catch (InvalidFileException ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                ClearAliasListPath();
            }
            catch (FileNotFoundException ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                ClearAliasListPath();
            }
            catch (WorksheetNotFoundException ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                ClearAliasListPath();
            }
            catch (Exception ex)
            {
                MessageBox.Show("An unexpected error occurred: " + ex.Message, "Error", MessageBoxButton.OK,
                    MessageBoxImage.Error);
                ClearAliasListPath();
            }
            catch
            {
                MessageBox.Show("An unexpected error occurred", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                ClearAliasListPath();
            }
        }
    }

    private void RemoveAliasListExecute(object obj)
    {
        Console.WriteLine("Remove Alias Map");
        ClearAliasListPath();
    }

    private void GenerateMergeListExecute(object obj)
    {
        if (string.IsNullOrEmpty(EmployeesListFilePath))
        {
            MessageBox.Show("Please upload the Employees List before generating the Merge List (The excel file with all the aliases and the corresponding company emails).", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        if (string.IsNullOrEmpty(AliasListFilePath))
        {
            MessageBox.Show("Please upload the Alias List before generating the Merge List", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        try
        {
            DocumentMerger.MergeDocuments(Employees, AliasListWorksheet);
        }
        catch (Exception ex)
        {
            MessageBox.Show("An unexpected error occurred: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private string GetFileName(string filePath)
    {
        return Path.GetFileName(filePath);
    }
}