using Bytescout.Spreadsheet;
using EdufocalCertificateGenerator.Models;

namespace EdufocalCertificateGenerator.Services;

public class DocumentReader
{
    public Spreadsheet Document { get; set; }

    public DocumentReader(String path)
    {
        Spreadsheet document = new Spreadsheet();
        document.LoadFromFile(path);
        Document = document;
    }

    public void GenerateList(EmployeeMap employees)
    {
        Worksheet worksheet = Document.Workbook.Worksheets.ByName("Sheet1");

        for (int i = 1; i <= worksheet.UsedRangeRowMax; i++)
        {
            string aliasEmail = worksheet.Cell(i, 3).Value as string;
            string firstName = worksheet.Cell(i, 0).Value as string;
            string lastName = worksheet.Cell(i, 1).Value as string;
            string companyEmail = worksheet.Cell(i, 2).Value as string;

            employees.AddUser(aliasEmail, firstName, lastName, companyEmail);
        }
    }
}