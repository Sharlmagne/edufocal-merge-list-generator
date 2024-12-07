using System.IO;
using System.Windows;
using Bytescout.Spreadsheet;
using EdufocalMergeListGenerator.Models;
using Microsoft.Win32;

namespace EdufocalMergeListGenerator.Services;

public static class DocumentMerger
{
    private static readonly string _mailMergeTemplatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "boj_mail_merge_template.xlsx");

    public static void MergeDocuments(Dictionary<string, EmployeeInfo> employees, Worksheet aliasList)
    {
        var Today = DateTime.Now.ToString("yyyy-MM-dd");
        var dialog = SaveMergedDocument($"BOJ Employees Mail Merge List_{Today}");

        if (dialog.ShowDialog() != true)
        {
            return;
        }

        var newFilePath = dialog.FileName;

        // Clone the template file
        File.Copy(_mailMergeTemplatePath, newFilePath, true);

        // Merge documents
        var mailMergeTemplate = new DocumentReader(newFilePath);
        var mailMergeWorksheet = mailMergeTemplate.Document.Workbook.Worksheets.ByName("Sheet1");

        for (var i = 1; i <= aliasList.UsedRangeRowMax; i++)
        {
            string aliasEmail = aliasList.Cell(i, 0).Value as string;
            string courseName = aliasList.Cell(i, 1).Value as string;
            string moduleCompletion = aliasList.Cell(i, 2).Value as string;
            var dateValue = aliasList.Cell(i, 3).Value;

            // Create a DateTime object from the date value
            string dateAwarded = dateValue != null ? DateTime.FromOADate((double)dateValue).ToString("dd-MMM-yy") : null;


            Console.WriteLine("Alias Email: " + aliasEmail);
            Console.WriteLine("Course Name: " + courseName);
            Console.WriteLine("Module Completion: " + moduleCompletion);
            Console.WriteLine("Date Awarded: " + dateAwarded);

            if (aliasEmail != null && courseName != null && dateAwarded != null && moduleCompletion != null)
            {
                Console.WriteLine(aliasEmail);

                if (employees.TryGetValue(aliasEmail, out var employee))
                {
                    Console.WriteLine("Employee Found: " + employee.FirstName);
                    mailMergeWorksheet.Cell(i, 0).Value = employee.FirstName;
                    mailMergeWorksheet.Cell(i, 1).Value = employee.LastName;
                    mailMergeWorksheet.Cell(i, 2).Value = employee.CompanyEmail;
                    mailMergeWorksheet.Cell(i, 3).Value = aliasEmail;
                    mailMergeWorksheet.Cell(i, 4).Value = courseName;
                    mailMergeWorksheet.Cell(i, 5).Value = moduleCompletion;
                    mailMergeWorksheet.Cell(i, 6).Value = dateAwarded;
                }
            }
        }

        mailMergeTemplate.Document.SaveAs(newFilePath);

        MessageBox.Show("Mail Merge List generated successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private static SaveFileDialog SaveMergedDocument(string fileName = "MergedDocument")
    {
        return new SaveFileDialog()
        {
            Filter = "Excel Files|*.xlsx",
            FileName = fileName
        };
    }
}