using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Excel = DocumentFormat.OpenXml.Spreadsheet;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;

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
                case "import-styles":
                case "styles-import":  // Keep for backward compatibility
                case "replace-styles":  // Keep for backward compatibility
                    HandleReplaceStyles(args);
                    break;
                case "import-snippet-styles":
                case "styles-import-snippet":  // Keep for backward compatibility
                case "replace-styles-from-snippet":  // Keep for backward compatibility
                    HandleReplaceStylesFromSnippet(args);
                    break;
                case "test-clean":
                    HandleTestClean(args);
                    break;
                case "validate":
                    HandleValidate(args);
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
                Console.WriteLine("Types: xl-book, xl-mbook, xl-template, xl-mtemplate, word-doc, word-mdoc, word-template, word-mtemplate");
                return;
            }
            
            string templateType = args[1].ToLower();
            string fileName = args.Length > 2 ? args[2] : GetDefaultFileName(templateType);
            string outputPath = GetOutputPath(templateType, fileName);
            
            try
            {
                switch (templateType)
                {
                    case "xl-book":
                        CreateExcelBook(outputPath);
                        break;
                    case "xl-mbook":
                        CreateExcelMacroBook(outputPath);
                        break;
                    case "xl-template":
                        CreateExcelTemplate(outputPath);
                        break;
                    case "xl-mtemplate":
                        CreateExcelMacroTemplate(outputPath);
                        break;
                    case "word-doc":
                        CreateWordDoc(outputPath);
                        break;
                    case "word-mdoc":
                        CreateWordMacroDoc(outputPath);
                        break;
                    case "word-template":
                        CreateWordTemplate(outputPath);
                        break;
                    case "word-mtemplate":
                        CreateWordMacroTemplate(outputPath);
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
        
        static void CreateExcelTemplate(string outputPath)
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
        
        static void CreateExcelMacroTemplate(string outputPath)
        {
            using (var document = SpreadsheetDocument.Create(outputPath, SpreadsheetDocumentType.MacroEnabledTemplate))
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
        
        static void CreateExcelMacroBook(string outputPath)
        {
            using (var document = SpreadsheetDocument.Create(outputPath, SpreadsheetDocumentType.MacroEnabledWorkbook))
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
        
        static void CreateWordTemplate(string outputPath)
        {
            using (var document = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Template))
            {
                var mainPart = document.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                mainPart.Document.Save();
            }
        }
        
        static void CreateWordMacroTemplate(string outputPath)
        {
            using (var document = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.MacroEnabledTemplate))
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
        
        static void CreateWordMacroDoc(string outputPath)
        {
            using (var document = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.MacroEnabledDocument))
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
                "xl-template" => "Book",
                "xl-mtemplate" => "Book",
                "xl-book" => "Book",
                "xl-mbook" => "Book",
                "word-template" => "Doc",
                "word-mtemplate" => "Doc",
                "word-doc" => "Document",
                "word-mdoc" => "Document",
                _ => "Template"
            };
        }
        
        static string GetOutputPath(string templateType, string fileName)
        {
            string extension = templateType switch
            {
                "xl-template" => ".xltx",
                "xl-mtemplate" => ".xltm",
                "xl-book" => ".xlsx",
                "xl-mbook" => ".xlsm",
                "word-template" => ".dotx",
                "word-mtemplate" => ".dotm",
                "word-doc" => ".docx",
                "word-mdoc" => ".docm",
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
        
        static void HandleReplaceStyles(string[] args)
        {
            if (args.Length < 3)
            {
                Console.WriteLine("Usage: import-styles <target-doc> <source-doc>");
                Console.WriteLine("  target-doc: Document that will receive new styles (will be modified)");
                Console.WriteLine("  source-doc: Document with styles to copy from (read-only)");
                return;
            }

            string targetPath = args[1];
            string sourcePath = args[2];

            if (!File.Exists(targetPath))
            {
                Console.WriteLine($"Target document not found: {targetPath}");
                return;
            }

            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source document not found: {sourcePath}");
                return;
            }

            try
            {
                using (WordprocessingDocument targetDoc = WordprocessingDocument.Open(targetPath, true))
                {
                    using (WordprocessingDocument sourceDoc = WordprocessingDocument.Open(sourcePath, false))
                    {
                        // Get the StyleDefinitionsPart from both documents
                        StyleDefinitionsPart targetStylesPart = targetDoc.MainDocumentPart?.StyleDefinitionsPart;
                        StyleDefinitionsPart sourceStylesPart = sourceDoc.MainDocumentPart?.StyleDefinitionsPart;

                        if (sourceStylesPart?.Styles == null)
                        {
                            Console.WriteLine("Source document does not have a styles part.");
                            return;
                        }

                        if (targetStylesPart == null)
                        {
                            // Target doesn't have styles part, create one
                            targetStylesPart = targetDoc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
                            Console.WriteLine("Created new styles part in target document.");
                        }

                        // Clone and replace the entire Styles element
                        targetStylesPart.Styles = (Styles)sourceStylesPart.Styles.CloneNode(true);
                        targetStylesPart.Styles.Save();

                        Console.WriteLine($"Successfully replaced styles in: {targetPath}");
                        Console.WriteLine($"Styles copied from: {sourcePath}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error replacing styles: {ex.Message}");
            }
        }

        static void HandleReplaceStylesFromSnippet(string[] args)
        {
            if (args.Length < 3)
            {
                Console.WriteLine("Usage: import-snippet-styles <target-doc> <snippet-file> <snippet-id>");
                Console.WriteLine("  target-doc: Document that will receive styles (will be modified)");
                Console.WriteLine("  snippet-file: XML file containing style snippets");
                Console.WriteLine("  snippet-id: ID of the snippet to use (e.g., 'listStylesDefault')");
                Console.WriteLine("");
                Console.WriteLine("Example:");
                Console.WriteLine("  import-snippet-styles target.docx core/partials/styles/list-styles.xml listStylesDefault");
                return;
            }

            string targetPath = args[1];
            string snippetFile = args[2];
            string snippetId = args.Length > 3 ? args[3] : null;

            if (!File.Exists(targetPath))
            {
                Console.WriteLine($"Target document not found: {targetPath}");
                return;
            }

            if (!File.Exists(snippetFile))
            {
                Console.WriteLine($"Snippet file not found: {snippetFile}");
                return;
            }

            try
            {
                // Load snippet XML
                XDocument snippetDoc = XDocument.Load(snippetFile);
                XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

                // Find the snippet element
                XElement snippetElement = null;
                if (string.IsNullOrEmpty(snippetId))
                {
                    // Use the root snippet if no ID specified
                    snippetElement = snippetDoc.Root;
                }
                else
                {
                    // Check if root element has the matching ID
                    if (snippetDoc.Root?.Attribute("id")?.Value == snippetId)
                    {
                        snippetElement = snippetDoc.Root;
                    }
                    else
                    {
                        // Find snippet by ID in descendants
                        snippetElement = snippetDoc.Descendants("snippet")
                            .FirstOrDefault(s => s.Attribute("id")?.Value == snippetId);
                    }
                    
                    if (snippetElement == null)
                    {
                        Console.WriteLine($"Snippet with id '{snippetId}' not found in {snippetFile}");
                        Console.WriteLine($"Available snippet IDs:");
                        if (snippetDoc.Root?.Attribute("id") != null)
                        {
                            Console.WriteLine($"  - {snippetDoc.Root.Attribute("id").Value}");
                        }
                        foreach (var s in snippetDoc.Descendants("snippet"))
                        {
                            var id = s.Attribute("id")?.Value;
                            if (id != null)
                                Console.WriteLine($"  - {id}");
                        }
                        return;
                    }
                }

                // Extract style elements from snippet
                var snippetStyles = snippetElement.Elements(w + "style").ToList();
                
                if (snippetStyles.Count == 0)
                {
                    Console.WriteLine("No styles found in snippet.");
                    return;
                }

                Console.WriteLine($"Found {snippetStyles.Count} style(s) in snippet '{snippetId ?? "default"}'");

                using (WordprocessingDocument targetDoc = WordprocessingDocument.Open(targetPath, true))
                {
                    StyleDefinitionsPart targetStylesPart = targetDoc.MainDocumentPart?.StyleDefinitionsPart;

                    if (targetStylesPart == null)
                    {
                        targetStylesPart = targetDoc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
                        targetStylesPart.Styles = new Styles();
                        Console.WriteLine("Created new styles part in target document.");
                    }

                    if (targetStylesPart.Styles == null)
                    {
                        targetStylesPart.Styles = new Styles();
                    }

                    var targetStyles = targetStylesPart.Styles;
                    int replacedCount = 0;
                    int addedCount = 0;

                    // Process each style from snippet
                    foreach (var snippetStyleElement in snippetStyles)
                    {
                        string styleId = snippetStyleElement.Attribute(w + "styleId")?.Value;
                        
                        if (string.IsNullOrEmpty(styleId))
                        {
                            Console.WriteLine("Warning: Skipping style without styleId");
                            continue;
                        }

                        // Find existing style in target by styleId
                        var existingStyle = targetStyles.Elements<Style>()
                            .FirstOrDefault(s => s.StyleId?.Value == styleId);

                        // Convert XElement to OpenXml Style
                        string styleXml = snippetStyleElement.ToString();
                        using (var stringReader = new StringReader(styleXml))
                        using (var reader = System.Xml.XmlReader.Create(stringReader))
                        {
                            reader.MoveToContent();
                            var newStyle = new Style(reader.ReadOuterXml());

                            if (existingStyle != null)
                            {
                                // Remove old style
                                existingStyle.Remove();
                                replacedCount++;
                            }
                            else
                            {
                                addedCount++;
                            }

                            // Add new style at the end (will reorder later)
                            targetStyles.AppendChild(newStyle);
                        }
                    }

                    // Reorder styles to match snippet order
                    // Get all styles that came from snippet
                    var snippetStyleIds = snippetStyles
                        .Select(s => s.Attribute(w + "styleId")?.Value)
                        .Where(id => !string.IsNullOrEmpty(id))
                        .ToList();

                    // Get all current styles
                    var allStyles = targetStyles.Elements<Style>().ToList();
                    
                    // Separate snippet styles from other styles
                    var snippetStylesInDoc = allStyles
                        .Where(s => snippetStyleIds.Contains(s.StyleId?.Value))
                        .OrderBy(s => snippetStyleIds.IndexOf(s.StyleId?.Value))
                        .ToList();
                    
                    var otherStyles = allStyles
                        .Where(s => !snippetStyleIds.Contains(s.StyleId?.Value))
                        .ToList();

                    // Clear and rebuild in correct order
                    targetStyles.RemoveAllChildren<Style>();
                    
                    // Add other styles first (maintain their original position)
                    foreach (var style in otherStyles)
                    {
                        targetStyles.AppendChild(style);
                    }
                    
                    // Then add snippet styles in snippet order
                    foreach (var style in snippetStylesInDoc)
                    {
                        targetStyles.AppendChild(style);
                    }

                    targetStylesPart.Styles.Save();

                    Console.WriteLine($"Successfully processed {replacedCount} replaced, {addedCount} added style(s)");
                    Console.WriteLine($"Styles from snippet '{snippetId ?? "default"}' applied to: {targetPath}");
                    Console.WriteLine($"Style order preserved from snippet.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error replacing styles from snippet: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
            }
        }
        
        static void ShowUsage()
        {
            Console.WriteLine("Usage:");
            Console.WriteLine("  dotnet run pack <template-path>");
            Console.WriteLine("  dotnet run unpack <template-path>");
            Console.WriteLine("  dotnet run create <type> [name]");
            Console.WriteLine("  dotnet run import-styles <target-doc> <source-doc>");
            Console.WriteLine("  dotnet run import-snippet-styles <target-doc> <snippet-file> [snippet-id]");
            Console.WriteLine("  dotnet run validate <file-path>");
            Console.WriteLine("  dotnet run test-clean [--tmp] [--out]");
            Console.WriteLine("");
            Console.WriteLine("Create types:");
            Console.WriteLine("  xl-template           - Create Excel Book template (.xltx)");
            Console.WriteLine("  xl-mtemplate          - Create Excel Macro-Enabled Book template (.xltm)");
            Console.WriteLine("  xl-book               - Create Excel workbook (.xlsx)");
            Console.WriteLine("  xl-mbook              - Create Excel Macro-Enabled workbook (.xlsm)");
            Console.WriteLine("  word-template         - Create Word document template (.dotx)");
            Console.WriteLine("  word-mtemplate        - Create Word Macro-Enabled document template (.dotm)");
            Console.WriteLine("  word-doc              - Create Word document (.docx)");
            Console.WriteLine("  word-mdoc             - Create Word Macro-Enabled document (.docm)");
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

        static void HandleValidate(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Error: File path required");
                Console.WriteLine("Usage: validate <file-path>");
                return;
            }

            string filePath = args[1];
            
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: File not found: {filePath}");
                return;
            }

            string fileDir = Path.GetDirectoryName(Path.GetFullPath(filePath)) ?? Directory.GetCurrentDirectory();
            string reportsDir = Path.Combine(fileDir, "reports");
            Directory.CreateDirectory(reportsDir);

            string fileName = Path.GetFileNameWithoutExtension(filePath);
            string reportPath = Path.Combine(reportsDir, $"{fileName}-report.txt");

            try
            {
                Console.WriteLine($"Validating: {filePath}");
                
                using var writer = new StreamWriter(reportPath);
                writer.WriteLine("OOXML Validation Report");
                writer.WriteLine("======================");
                writer.WriteLine($"File: {Path.GetFileName(filePath)}");
                writer.WriteLine($"Full Path: {Path.GetFullPath(filePath)}");
                writer.WriteLine($"Timestamp: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                writer.WriteLine($"File Size: {new FileInfo(filePath).Length:N0} bytes");
                writer.WriteLine();

                string ext = Path.GetExtension(filePath).ToLower();
                bool isValid = true;
                int errorCount = 0;

                switch (ext)
                {
                    case ".docx":
                    case ".docm":
                    case ".dotx":
                    case ".dotm":
                        errorCount = ValidateWordDocument(filePath, writer);
                        break;
                    case ".xlsx":
                    case ".xlsm":
                    case ".xltx":
                    case ".xltm":
                        errorCount = ValidateExcelDocument(filePath, writer);
                        break;
                    case ".pptx":
                    case ".pptm":
                    case ".potx":
                    case ".potm":
                        errorCount = ValidatePowerPointDocument(filePath, writer);
                        break;
                    default:
                        writer.WriteLine($"ERROR: Unsupported file type: {ext}");
                        writer.WriteLine("Supported types: .docx, .docm, .dotx, .dotm, .xlsx, .xlsm, .xltx, .xltm, .pptx, .pptm, .potx, .potm");
                        isValid = false;
                        errorCount = 1;
                        break;
                }

                writer.WriteLine();
                writer.WriteLine("======================");
                writer.WriteLine($"Validation {(errorCount == 0 ? "PASSED" : "FAILED")}");
                writer.WriteLine($"Total Errors: {errorCount}");

                Console.WriteLine($"Validation {(errorCount == 0 ? "PASSED" : "FAILED")} - {errorCount} error(s) found");
                Console.WriteLine($"Report saved to: {reportPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during validation: {ex.Message}");
                using var writer = File.AppendText(reportPath);
                writer.WriteLine();
                writer.WriteLine("CRITICAL ERROR:");
                writer.WriteLine(ex.ToString());
            }
        }

        static int ValidateWordDocument(string filePath, StreamWriter writer)
        {
            writer.WriteLine("Document Type: Word");
            writer.WriteLine();

            try
            {
                using var doc = WordprocessingDocument.Open(filePath, false);
                var validator = new DocumentFormat.OpenXml.Validation.OpenXmlValidator();
                var errors = validator.Validate(doc).ToList();

                if (errors.Count == 0)
                {
                    writer.WriteLine("No validation errors found.");
                }
                else
                {
                    writer.WriteLine($"Found {errors.Count} validation error(s):");
                    writer.WriteLine();

                    int errorNum = 1;
                    foreach (var error in errors)
                    {
                        writer.WriteLine($"Error {errorNum++}:");
                        writer.WriteLine($"  Description: {error.Description}");
                        writer.WriteLine($"  Error Type: {error.ErrorType}");
                        writer.WriteLine($"  Part: {error.Part?.Uri}");
                        writer.WriteLine($"  Path: {error.Path?.XPath}");
                        if (error.RelatedNode != null)
                        {
                            writer.WriteLine($"  Node: {error.RelatedNode.LocalName}");
                        }
                        writer.WriteLine();
                    }
                }

                return errors.Count;
            }
            catch (Exception ex)
            {
                writer.WriteLine($"ERROR: Failed to validate Word document: {ex.Message}");
                return 1;
            }
        }

        static int ValidateExcelDocument(string filePath, StreamWriter writer)
        {
            writer.WriteLine("Document Type: Excel");
            writer.WriteLine();

            try
            {
                using var doc = SpreadsheetDocument.Open(filePath, false);
                var validator = new DocumentFormat.OpenXml.Validation.OpenXmlValidator();
                var errors = validator.Validate(doc).ToList();

                if (errors.Count == 0)
                {
                    writer.WriteLine("No validation errors found.");
                }
                else
                {
                    writer.WriteLine($"Found {errors.Count} validation error(s):");
                    writer.WriteLine();

                    int errorNum = 1;
                    foreach (var error in errors)
                    {
                        writer.WriteLine($"Error {errorNum++}:");
                        writer.WriteLine($"  Description: {error.Description}");
                        writer.WriteLine($"  Error Type: {error.ErrorType}");
                        writer.WriteLine($"  Part: {error.Part?.Uri}");
                        writer.WriteLine($"  Path: {error.Path?.XPath}");
                        if (error.RelatedNode != null)
                        {
                            writer.WriteLine($"  Node: {error.RelatedNode.LocalName}");
                        }
                        writer.WriteLine();
                    }
                }

                return errors.Count;
            }
            catch (Exception ex)
            {
                writer.WriteLine($"ERROR: Failed to validate Excel document: {ex.Message}");
                return 1;
            }
        }

        static int ValidatePowerPointDocument(string filePath, StreamWriter writer)
        {
            writer.WriteLine("Document Type: PowerPoint");
            writer.WriteLine();

            try
            {
                using var doc = PresentationDocument.Open(filePath, false);
                var validator = new DocumentFormat.OpenXml.Validation.OpenXmlValidator();
                var errors = validator.Validate(doc).ToList();

                if (errors.Count == 0)
                {
                    writer.WriteLine("No validation errors found.");
                }
                else
                {
                    writer.WriteLine($"Found {errors.Count} validation error(s):");
                    writer.WriteLine();

                    int errorNum = 1;
                    foreach (var error in errors)
                    {
                        writer.WriteLine($"Error {errorNum++}:");
                        writer.WriteLine($"  Description: {error.Description}");
                        writer.WriteLine($"  Error Type: {error.ErrorType}");
                        writer.WriteLine($"  Part: {error.Part?.Uri}");
                        writer.WriteLine($"  Path: {error.Path?.XPath}");
                        if (error.RelatedNode != null)
                        {
                            writer.WriteLine($"  Node: {error.RelatedNode.LocalName}");
                        }
                        writer.WriteLine();
                    }
                }

                return errors.Count;
            }
            catch (Exception ex)
            {
                writer.WriteLine($"ERROR: Failed to validate PowerPoint document: {ex.Message}");
                return 1;
            }
        }
    }
}
