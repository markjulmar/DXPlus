// See https://aka.ms/new-console-template for more information

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

ValidateWordDocument(@"C:\users\mark\onedrive\desktop\test.docx");


void ValidateWordDocument(string filepath)
{
    using var wordprocessingDocument = WordprocessingDocument.Open(filepath, true);
    try
    {
        OpenXmlValidator validator = new OpenXmlValidator();
        int count = 0;
        foreach (ValidationErrorInfo error in
                 validator.Validate(wordprocessingDocument))
        {
            count++;
            Console.WriteLine("Error " + count);
            Console.WriteLine("Description: " + error.Description);
            Console.WriteLine("ErrorType: " + error.ErrorType);
            Console.WriteLine("Node: " + error.Node);
            Console.WriteLine("Path: " + error.Path.XPath);
            Console.WriteLine("Part: " + error.Part.Uri);
            Console.WriteLine("-------------------------------------------");
        }

        Console.WriteLine("count={0}", count);
    }

    catch (Exception ex)
    {
        Console.WriteLine(ex.Message);
    }

    wordprocessingDocument.Close();
}