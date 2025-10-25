using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Excel = DocumentFormat.OpenXml.Spreadsheet;
using System.IO.Compression;
using System.Linq;

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
                case "test-clean":
                    HandleTestClean(args);
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
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: pack <expanded-folder> <output-file>");
                return;
            }

            string sourceDir = args[1];
            string outputFile = args.Length > 2 ? args[2] : "";

            if (!Directory.Exists(sourceDir))
            {
                Console.WriteLine($"Source directory not found: {sourceDir}");
                return;
            }

            if (string.IsNullOrEmpty(outputFile))
            {
                // Try to write to the standard out folder when source follows builds/{agency}/{category}/{template}/in
                try
                {
                    var srcDirInfo = new DirectoryInfo(sourceDir);
                    // Look for a parent path that ends with /in
                    if (srcDirInfo.Name.Equals("in", StringComparison.OrdinalIgnoreCase) && srcDirInfo.Parent != null)
                    {
                        var templateDir = srcDirInfo.Parent; // .../{template}
                        var outDir = Path.Combine(templateDir.Parent?.FullName ?? templateDir.FullName, "out");
                        Directory.CreateDirectory(outDir);

                        // Determine filename similar to earlier alias logic
                        var templateName = templateDir.Name;
                        string ext = ".dotx";
                        string filename = templateName + ext;
                        if (templateName.Contains("Normal", StringComparison.OrdinalIgnoreCase))
                        {
                            ext = ".dotm";
                            filename = "Normal" + ext;
                        }
                        else if (templateName.Contains("Sheet", StringComparison.OrdinalIgnoreCase))
                        {
                            ext = ".xltx";
                            filename = "Sheet" + ext;
                        }
                        else if (templateName.Contains("Book", StringComparison.OrdinalIgnoreCase))
                        {
                            ext = ".xltx";
                            filename = "Book" + ext;
                        }
                        else if (templateName.Contains("Theme", StringComparison.OrdinalIgnoreCase))
                        {
                            ext = ".thmx";
                            filename = templateName + ext;
                        }
                        else if (templateName.Contains("Letterhead", StringComparison.OrdinalIgnoreCase) || templateName.Contains("Memo", StringComparison.OrdinalIgnoreCase) || templateName.Contains("Form", StringComparison.OrdinalIgnoreCase))
                        {
                            ext = ".dotx";
                            filename = templateName + ext;
                        }

                        outputFile = Path.Combine(outDir, filename);
                        Console.WriteLine($"Auto-generating output in out folder: {outputFile}");
                    }
                    else
                    {
                        // fallback to temp
                        var name = srcDirInfo.Name ?? "packed";
                        var safeName = name.Replace(Path.DirectorySeparatorChar, '_').Replace(Path.AltDirectorySeparatorChar, '_');
                        var ts = DateTime.UtcNow.ToString("yyyyMMddHHmmss");
                        var tempPath = Path.GetTempPath();
                        outputFile = Path.Combine(tempPath, $"{safeName}_{ts}.packed");
                        Console.WriteLine($"No output specified — writing pack to temporary file: {outputFile}");
                    }
                }
                catch
                {
                    var name = new DirectoryInfo(sourceDir).Name ?? "packed";
                    var safeName = name.Replace(Path.DirectorySeparatorChar, '_').Replace(Path.AltDirectorySeparatorChar, '_');
                    var ts = DateTime.UtcNow.ToString("yyyyMMddHHmmss");
                    var tempPath = Path.GetTempPath();
                    outputFile = Path.Combine(tempPath, $"{safeName}_{ts}.packed");
                    Console.WriteLine($"No output specified — writing pack to temporary file: {outputFile}");
                }
            }

            try
            {
                if (File.Exists(outputFile)) File.Delete(outputFile);
                ZipFile.CreateFromDirectory(sourceDir, outputFile, CompressionLevel.Optimal, false);
                Console.WriteLine($"Packed: {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error packing: {ex.Message}");
            }
        }
        
        static void HandleUnpack(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: unpack <template-file> [output-folder]");
                return;
            }

            string templateFile = args[1];
            string outputFolder = args.Length > 2 ? args[2] : "";

            if (!File.Exists(templateFile))
            {
                Console.WriteLine($"Template file not found: {templateFile}");
                return;
            }

            if (string.IsNullOrEmpty(outputFolder))
            {
                // Default: create expanded folder beside the template file
                var baseName = Path.GetFileNameWithoutExtension(templateFile);
                var parentDir = Path.GetDirectoryName(templateFile) ?? ".";
                outputFolder = Path.Combine(parentDir, baseName + "_expanded");

                // If folder already exists, append timestamp to avoid removing user data
                if (Directory.Exists(outputFolder))
                {
                    var ts = DateTime.UtcNow.ToString("yyyyMMddHHmmss");
                    outputFolder = Path.Combine(parentDir, baseName + "_expanded_" + ts);
                }

                Console.WriteLine($"No output specified — unpacking to: {outputFolder}");
            }

            try
            {
                if (Directory.Exists(outputFolder))
                {
                    Console.WriteLine($"Output folder already exists, removing: {outputFolder}");
                    Directory.Delete(outputFolder, true);
                }

                Directory.CreateDirectory(outputFolder);
                ZipFile.ExtractToDirectory(templateFile, outputFolder);
                Console.WriteLine($"Unpacked to: {outputFolder}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error unpacking: {ex.Message}");
            }
        }
        
        static void ShowUsage()
        {
            Console.WriteLine("Usage:");
            Console.WriteLine("  dotnet run pack <template-path>");
            Console.WriteLine("  dotnet run unpack <template-path>");
            Console.WriteLine("  dotnet run create <type> [name]");
            Console.WriteLine("  dotnet run test-clean [--tmp] [--out]");
            Console.WriteLine("");
            Console.WriteLine("Create types:");
            Console.WriteLine("  excel-book-template   - Create Excel Book template (.xltx)");
            Console.WriteLine("  excel-sheet-template  - Create Excel Sheet template (.xltx)");
            Console.WriteLine("  excel-book           - Create Excel workbook (.xlsx)");
            Console.WriteLine("  word-doc-template    - Create Word document template (.dotx)");
            Console.WriteLine("  word-doc             - Create Word document (.docx)");
        }

        static void HandleTestClean(string[] args)
        {
            bool cleanTmp = false;
            bool cleanOut = false;

            foreach (var a in args.Skip(1))
            {
                if (a == "--tmp") cleanTmp = true;
                if (a == "--out") cleanOut = true;
            }

            // If no flags provided, clean both
            if (!cleanTmp && !cleanOut)
            {
                cleanTmp = true;
                cleanOut = true;
            }

            if (cleanTmp)
            {
                var tmp = Path.GetTempPath();
                Console.WriteLine($"Cleaning temp folder: {tmp}");
                try
                {
                    var dirs = Directory.GetDirectories(tmp, "*_expanded_*", SearchOption.TopDirectoryOnly)
                        .Concat(Directory.GetDirectories(tmp, "*_expanded", SearchOption.TopDirectoryOnly))
                        .ToArray();
                    var files = Directory.GetFiles(tmp, "*.packed", SearchOption.TopDirectoryOnly);

                    foreach (var d in dirs) { Console.WriteLine($"Removing: {d}"); Directory.Delete(d, true); }
                    foreach (var f in files) { Console.WriteLine($"Removing: {f}"); File.Delete(f); }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error cleaning tmp: {ex.Message}");
                }
            }

            if (cleanOut)
            {
                // Restrict search to builds/ to avoid accidental deletions
                var root = Path.Combine(Directory.GetCurrentDirectory(), "builds");
                if (!Directory.Exists(root))
                {
                    Console.WriteLine("No builds/ directory found, skipping out cleanup");
                    return;
                }

                Console.WriteLine($"Cleaning test artifacts under: {root}");
                try
                {
                    // Remove expanded folders and .packed files under builds/
                    var dirs = Directory.GetDirectories(root, "*_expanded_*", SearchOption.AllDirectories)
                        .Concat(Directory.GetDirectories(root, "*_expanded", SearchOption.AllDirectories))
                        .ToArray();
                    var files = Directory.GetFiles(root, "*.packed", SearchOption.AllDirectories);

                    foreach (var d in dirs) { Console.WriteLine($"Removing: {d}"); Directory.Delete(d, true); }
                    foreach (var f in files) { Console.WriteLine($"Removing: {f}"); File.Delete(f); }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error cleaning builds/: {ex.Message}");
                }
            }
        }
    }
}
