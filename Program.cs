using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Properties;

class Program
{
    static void Main(string[] args)
    {
        string folderPath = @"C:\Documents\Test 1";
        List<string> documentProperties = new List<string>();

        // Get all the Word documents in the specified folder and subfolders
        string[] docFiles = Directory.GetFiles(folderPath, "*.docx", SearchOption.AllDirectories);

        // Extract document properties and cross-reference
        foreach (string docFile in docFiles)
        {
            Document doc = new Document(docFile);

            // Extract document properties from the current document
            foreach (DocumentProperty property in doc.BuiltInDocumentProperties)
            {
                if (!documentProperties.Contains(property.Name))
                {
                    documentProperties.Add(property.Name);
                }
            }

            // Cross-reference document properties used in the current document
            Console.WriteLine($"Document: {Path.GetFileName(docFile)}");
            Console.WriteLine("Used Document Properties:");
            foreach (string propertyName in documentProperties)
            {
                string propertyValue = doc.BuiltInDocumentProperties[propertyName]?.ToString();
                if (!string.IsNullOrEmpty(propertyValue))
                {
                    Console.WriteLine($"  {propertyName}: {propertyValue}");
                }
            }
            Console.WriteLine();
        }

        // Display the cross-reference of document properties
        Console.WriteLine("Document Property Cross-Reference:");
        foreach (string propertyName in documentProperties)
        {
            Console.WriteLine($"Property: {propertyName}");
            Console.WriteLine("Used in Documents:");
            foreach (string docFile in docFiles)
            {
                Document doc = new Document(docFile);
                string propertyValue = doc.BuiltInDocumentProperties[propertyName]?.ToString();
                if (!string.IsNullOrEmpty(propertyValue))
                {
                    Console.WriteLine($"  {Path.GetFileName(docFile)}");
                }
            }
            Console.WriteLine();
        }
    }
}