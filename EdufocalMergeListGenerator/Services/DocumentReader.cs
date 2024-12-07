using System.IO;
using Bytescout.Spreadsheet;
using EdufocalMergeListGenerator.Exceptions;
using EdufocalMergeListGenerator.Models;
using FileNotFoundException = System.IO.FileNotFoundException;

namespace EdufocalMergeListGenerator.Services;

public class DocumentReader
{
    public Spreadsheet Document { get; set; }

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

    public void GenerateEmployeesList(EmployeesList employees)
    {
        ValidateWorksheet();

        Worksheet worksheet = Document.Workbook.Worksheets.ByName("Sheet1");

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

    private void ValidateWorksheet()
    {
        if (string.IsNullOrEmpty(Document?.Workbook?.Worksheets?.ByName("Sheet1")?.Name))
        {
            throw new WorksheetNotFoundException("Document does not contain a worksheet named 'Sheet1'");
        }
    }

    public void ValidateAliasList()
    {
        ValidateWorksheet();

        Worksheet worksheet = Document.Workbook.Worksheets.ByName("Sheet1");


        if (worksheet.UsedRangeRowMax < 1)
        {
            throw new InvalidFileException("The worksheet does not contain any data. Please ensure the worksheet contains data before proceeding.");
        }
        else if (
            worksheet.Cell(0, 0).Value as string != "Alias_Email" ||
            worksheet.Cell(0, 1).Value as string != "Course_Name" ||
            worksheet.Cell(0, 2).Value as string != "Module_Completion" ||
            worksheet.Cell(0, 3).Value as string != "Date_Awarded" ||
            worksheet.UsedRangeColumnMax < 3
            )
        {
            throw new InvalidFileException("The worksheet does not contain the required columns. Please ensure the worksheet contains the following columns (In the exact order): 'Alias_Email', 'Course_Name', 'Module_Completion', 'Date_Awarded'");
        }
    }
}