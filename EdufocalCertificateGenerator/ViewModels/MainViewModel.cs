using System.IO;
using System.Windows.Input;
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

    public string MapFilePath
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

    public MainViewModel()
    {
        UploadUserMapCommand = new RelayCommand(UploadUserMapExecute, CanUploadUserMap);
        GenerateCertificateCommand = new RelayCommand(GenerateCertificateExecute, CanGenerateCertificate);
    }

    private void UploadUserMapExecute(object obj)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls|CSV files (*.csv)|*.csv"
        };

        if (dialog.ShowDialog() == true)
        {
            MapFilePath = dialog.FileName;
            FileName = GetFileName(MapFilePath);
        }
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