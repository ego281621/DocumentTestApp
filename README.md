# Document Property Cross-Reference Tool

This tool generates a cross-reference for document properties in Microsoft Word documents. It scans a folder and its subfolders for Word documents (`.docx` files), extracts document properties, and provides a cross-reference of used document properties for each document and a list of documents where each property is used.

## Requirements

- [.NET Core](https://dotnet.microsoft.com/download) installed.
- Aspose.Words library (version 23.8.0) [NuGet package](https://www.nuget.org/packages/Aspose.Words/23.8.0) installed.

## Usage

1. Clone this repository or download the source code.
2. Open a terminal or command prompt and navigate to the project directory.
3. Run the tool by executing the following command:

   ```bash
   dotnet run --project DocumentPropertyCrossReference

Replace DocumentPropertyCrossReference with the actual path to the project folder if needed.

The tool will prompt you to provide the path to the folder containing Word documents you want to cross-reference.

The cross-reference results will be displayed in the terminal.

## Features
- Scans a folder and its subfolders for Word documents (.docx files).
- Extracts document properties from each document.
- Generates a cross-reference of used document properties for each document.
- Lists documents where each property is used.
  
## Sample Output
```
Document: SampleDocument.docx
Used Document Properties:
Title: Sample Document Title
Author: John Doe

Document Property Cross-Reference:
Property: Title
Used in Documents:
  SampleDocument.docx
  AnotherDocument.docx

Property: Author
Used in Documents:
  SampleDocument.docx
  YetAnotherDocument.docx

##License
This tool is licensed under the MIT License.

##Credits
This tool utilizes the Aspose.Words library to handle Word documents. Visit the Aspose.Words website for more information.
