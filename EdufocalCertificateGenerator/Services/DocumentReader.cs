using System.IO;
using Bytescout.Spreadsheet;
using EdufocalCertificateGenerator.Exceptions;
using EdufocalCertificateGenerator.Models;
using FileNotFoundException = System.IO.FileNotFoundException;

namespace EdufocalCertificateGenerator.Services;

public class DocumentReader
{
    private Spreadsheet Document { get; set; }

    public DocumentReader(string path)
    {
        if (!File.Exists(path))
        {
            throw new FileNotFoundException("File does not exist: " + path);
        }

        var document = new Spreadsheet();
        document.LoadFromFile(path);
        Document = document;
    }

    public void GenerateList(EmployeeMap employees)
    {
        if (string.IsNullOrEmpty(Document?.Workbook?.Worksheets?.ByName("Sheet1")?.Name))
        {
            throw new WorksheetNotFoundException("Document does not contain a worksheet named 'Sheet1'");
        }

        Worksheet worksheet = Document.Workbook.Worksheets.ByName("Sheet1");

        Console.WriteLine("First Row: " + worksheet.Cell(0, 0).Value );
        Console.WriteLine("Second Row: " + worksheet.Cell(0, 1).Value);
        Console.WriteLine("Third Row: " + worksheet.Cell(0, 2).Value);
        Console.WriteLine("Fourth Row: " + worksheet.Cell(0, 3).Value);
        Console.WriteLine("Row Count: " + worksheet.UsedRangeRowMax);
        Console.WriteLine("Column Count: " + worksheet.UsedRangeColumnMax);

        // Validate the worksheet
        if (worksheet.UsedRangeRowMax < 1)
        {
            throw new InvalidFileException("The worksheet does not contain any data. Please ensure the worksheet contains data before proceeding.");
        }
        else if (
            worksheet.Cell(0, 0).Value as string != "First_Name" ||
            worksheet.Cell(0, 1).Value as string != "Last_Name" ||
            worksheet.Cell(0, 2).Value as string != "Company_Email" ||
            worksheet.Cell(0, 3).Value as string != "Alias_Email" ||
            worksheet.UsedRangeColumnMax < 3
            )
        {
            throw new InvalidFileException("The worksheet does not contain the required columns. Please ensure the worksheet contains the following columns (In the exact order): 'First_Name', 'Last_Name', 'Company_Email', 'Alias_Email'");
        }

        for (int i = 1; i <= worksheet.UsedRangeRowMax; i++)
        {
            string aliasEmail = worksheet.Cell(i, 3).Value as string;
            string firstName = worksheet.Cell(i, 0).Value as string;
            string lastName = worksheet.Cell(i, 1).Value as string;
            string companyEmail = worksheet.Cell(i, 2).Value as string;

            if (aliasEmail != null && firstName != null && lastName != null && companyEmail != null)
            {
                employees.AddUser(aliasEmail, firstName, lastName, companyEmail);
            }
        }
    }
}