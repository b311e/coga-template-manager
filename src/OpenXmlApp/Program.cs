using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Excel = DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlApp
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("OpenXML Template Manager");
            
            if (args.Length == 0)
            {
                ShowUsage();
                return;
            }
            
            string command = args[0].ToLower();
            
            switch (command)
            {
                case "pack":
                    HandlePack(args);
                    break;
                case "unpack":
                    HandleUnpack(args);
                    break;
                case "create":
                    HandleCreate(args);
                    break;
                default:
                    Console.WriteLine($"Unknown command: {command}");
                    ShowUsage();
                    break;
            }
        }
        
        static void HandleCreate(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Error: Template type required");
                Console.WriteLine("Usage: create <type> [name]");
                Console.WriteLine("Types: excel-book-template, excel-sheet-template, excel-book, word-doc-template, word-doc");
                return;
            }
            
            string templateType = args[1].ToLower();
            string fileName = args.Length > 2 ? args[2] : GetDefaultFileName(templateType);
            string outputPath = GetOutputPath(templateType, fileName);
            
            try
            {
                switch (templateType)
                {
                    case "excel-book-template":
                        CreateExcelBookTemplate(outputPath);
                        break;
                    case "excel-sheet-template":
                        CreateExcelSheetTemplate(outputPath);
                        break;
                    case "excel-book":
                        CreateExcelBook(outputPath);
                        break;
                    case "word-doc-template":
                        CreateWordDocTemplate(outputPath);
                        break;
                    case "word-doc":
                        CreateWordDoc(outputPath);
                        break;
                    default:
                        Console.WriteLine($"Unknown template type: {templateType}");
                        return;
                }
                
                Console.WriteLine($"Created: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating template: {ex.Message}");
            }
        }
        
        static void CreateExcelBookTemplate(string outputPath)
        {
            using (var document = SpreadsheetDocument.Create(outputPath, SpreadsheetDocumentType.Template))
            {
                var workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Excel.Workbook();
                
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Excel.Worksheet(new Excel.SheetData());
                
                var sheets = workbookPart.Workbook.AppendChild(new Excel.Sheets());
                sheets.Append(new Excel.Sheet() 
                { 
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "Sheet1"
                });
                
                workbookPart.Workbook.Save();
            }
        }
        
        static void CreateExcelSheetTemplate(string outputPath)
        {
            CreateExcelBookTemplate(outputPath);
        }
        
        static void CreateExcelBook(string outputPath)
        {
            using (var document = SpreadsheetDocument.Create(outputPath, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Excel.Workbook();
                
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Excel.Worksheet(new Excel.SheetData());
                
                var sheets = workbookPart.Workbook.AppendChild(new Excel.Sheets());
                sheets.Append(new Excel.Sheet() 
                { 
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "Sheet1"
                });
                
                workbookPart.Workbook.Save();
            }
        }
        
        static void CreateWordDocTemplate(string outputPath)
        {
            using (var document = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Template))
            {
                var mainPart = document.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                mainPart.Document.Save();
            }
        }
        
        static void CreateWordDoc(string outputPath)
        {
            using (var document = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document))
            {
                var mainPart = document.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                mainPart.Document.Save();
            }
        }
        
        static string GetDefaultFileName(string templateType)
        {
            return templateType switch
            {
                "excel-book-template" => "Book",
                "excel-sheet-template" => "Sheet", 
                "excel-book" => "Workbook",
                "word-doc-template" => "Normal",
                "word-doc" => "Document",
                _ => "Template"
            };
        }
        
        static string GetOutputPath(string templateType, string fileName)
        {
            string extension = templateType switch
            {
                "excel-book-template" => ".xltx",
                "excel-sheet-template" => ".xltx",
                "excel-book" => ".xlsx", 
                "word-doc-template" => ".dotx",
                "word-doc" => ".docx",
                _ => ".tmp"
            };
            
            return $"{fileName}{extension}";
        }
        
        static void HandlePack(string[] args)
        {
            Console.WriteLine("Pack command not implemented yet");
        }
        
        static void HandleUnpack(string[] args)
        {
            Console.WriteLine("Unpack command not implemented yet");
        }
        
        static void ShowUsage()
        {
            Console.WriteLine("Usage:");
            Console.WriteLine("  dotnet run pack <template-path>");
            Console.WriteLine("  dotnet run unpack <template-path>");
            Console.WriteLine("  dotnet run create <type> [name]");
            Console.WriteLine("");
            Console.WriteLine("Create types:");
            Console.WriteLine("  excel-book-template   - Create Excel Book template (.xltx)");
            Console.WriteLine("  excel-sheet-template  - Create Excel Sheet template (.xltx)");
            Console.WriteLine("  excel-book           - Create Excel workbook (.xlsx)");
            Console.WriteLine("  word-doc-template    - Create Word document template (.dotx)");
            Console.WriteLine("  word-doc             - Create Word document (.docx)");
        }
    }
}
