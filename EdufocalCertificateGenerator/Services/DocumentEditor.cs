using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Win32;

namespace EdufocalCertificateGenerator.Services;

public static class DocumentEditor
{
    private const string _employeeNameTag = "{employee_name}";
    private const string _courseTag = "{course_name}";
    private const string _moduleCompletionTag = "{module_completion}";
    private const string _dateAwardedTag = "{date_awarded}";


    public static void EditWordDocument(string templateFilePath, string employeeName, string courseName, string moduleCompletion, string dateAwarded)
    {
        var dialog = SaveFileDialog($"{employeeName}_{courseName}_Certificate");

        if (dialog.ShowDialog() != true)
        {
            return;
        }

        var newFilePath = dialog.FileName;

        // Clone the template file
        File.Copy(templateFilePath, newFilePath, true);

        // Open the new document for editing
        using var wordDoc = WordprocessingDocument.Open(newFilePath, true);

        var body = wordDoc.MainDocumentPart.Document.Body;

        foreach (var text in body.Descendants<Text>())
        {
            Console.WriteLine(text.Text);
            if (text.Text.Contains(_employeeNameTag))
            {
                text.Text = text.Text.Replace(_employeeNameTag, employeeName);
                Console.WriteLine("Found Employee Name");
            }
            if (text.Text.Contains(_courseTag))
            {
                text.Text = text.Text.Replace(_courseTag, courseName);
                Console.WriteLine("Found Course Name");
            }
            if (text.Text.Contains(_moduleCompletionTag))
            {
                text.Text = text.Text.Replace(_moduleCompletionTag, moduleCompletion);
                Console.WriteLine("Found Module Completion");
            }
            if (text.Text.Contains(_dateAwardedTag))
            {
                text.Text = text.Text.Replace(_dateAwardedTag, dateAwarded);
                Console.WriteLine("Found Date Awarded");
            }
        }

        wordDoc.MainDocumentPart.Document.Save();
    }

    private static SaveFileDialog SaveFileDialog(string fileName = "Certificate")
    {
        return new SaveFileDialog
        {
            Filter = "Word files (*.docx)|*.docx",
            FileName = fileName
        };
    }

}