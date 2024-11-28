using System.Configuration;
using System.IO;
using System.Windows.Input;
using EdufocalCertificateGenerator.Models;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Win32;
using NISInspectorApp.Core;

namespace EdufocalCertificateGenerator.ViewModels;

public class MainViewModel: ViewModel
{

    private string _mapFilePath;
    private string _fileName;

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
            AppConfig.Sections.Add("UserMap", new UserMap());
            AppConfig.Save(ConfigurationSaveMode.Modified);
        }
        else
        {
            var userMap = (UserMap)AppConfig.Sections["UserMap"];
            MapFilePath = userMap.MapFilePath;
            FileName = userMap.FileName;
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

            var userMap = (UserMap)AppConfig.Sections["UserMap"];
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
        var userMap = (UserMap)AppConfig.Sections["UserMap"];

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
        // Generate certificate
        throw new NotImplementedException();
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