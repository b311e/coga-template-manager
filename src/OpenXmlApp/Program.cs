using System.IO;
using System.IO.Compression;
using System.Xml.Linq;

class Program
{
    static void Main(string[] args)
    {
        if (args.Length < 1)
        {
            Console.WriteLine("Usage:");
            Console.WriteLine("  pack   <sourceDir> <out.dotm>");
            Console.WriteLine("  unpack <in.dotm> [outDir]");
            Console.WriteLine("    If outDir is omitted, creates 'expanded' folder next to input file");
            return;
        }

        switch (args[0])
        {
            case "pack":
                PackDotm(args[1], args[2]);
                break;
            case "unpack":
                if (args.Length >= 3)
                {
                    // Two arguments provided: unpack <input> <output>
                    Unpack(args[1], args[2]);
                }
                else
                {
                    // One argument provided: auto-create "expanded" folder next to source
                    UnpackAuto(args[1]);
                }
                break;
            default:
                Console.WriteLine("Unknown command.");
                break;
        }
    }

    static void PackDotm(string sourceDir, string outFile)
    {
        if (!Directory.Exists(sourceDir))
            throw new DirectoryNotFoundException(sourceDir);

        var types = Path.Combine(sourceDir, "[Content_Types].xml");
        if (!File.Exists(types))
            throw new InvalidDataException("Missing required OOXML Content_Types.xml file.");

        // Determine if this is a Word or Excel document by checking for key files
        var isWordDoc = Directory.Exists(Path.Combine(sourceDir, "word"));
        var isExcelDoc = Directory.Exists(Path.Combine(sourceDir, "xl"));

        if (!isWordDoc && !isExcelDoc)
            throw new InvalidDataException("Neither Word nor Excel document structure found.");

        // Validate content types based on document type
        XNamespace ct = "http://schemas.openxmlformats.org/package/2006/content-types";
        var x = XDocument.Load(types);
        
        if (isWordDoc)
        {
            var docXml = Path.Combine(sourceDir, "word", "document.xml");
            if (!File.Exists(docXml))
                throw new InvalidDataException("Missing required Word document.xml file.");
                
            bool valid = x.Root?.Elements(ct + "Override")
                .Any(o => (string?)o.Attribute("PartName") == "/word/document.xml" &&
                          ((string?)o.Attribute("ContentType"))?.Contains("word") == true) == true;
            if (!valid)
                throw new InvalidDataException("Invalid Word document content type.");
        }
        else if (isExcelDoc)
        {
            var workbookXml = Path.Combine(sourceDir, "xl", "workbook.xml");
            if (!File.Exists(workbookXml))
                throw new InvalidDataException("Missing required Excel workbook.xml file.");
                
            bool valid = x.Root?.Elements(ct + "Override")
                .Any(o => (string?)o.Attribute("PartName") == "/xl/workbook.xml" &&
                          ((string?)o.Attribute("ContentType"))?.Contains("spreadsheetml") == true) == true;
            if (!valid)
                throw new InvalidDataException("Invalid Excel document content type.");
        }

        if (File.Exists(outFile)) File.Delete(outFile);
        ZipFile.CreateFromDirectory(sourceDir, outFile, CompressionLevel.Optimal, includeBaseDirectory: false);

        Console.WriteLine($"Created {outFile}");
    }

    static void Unpack(string inFile, string outDir)
    {
        if (Directory.Exists(outDir)) Directory.Delete(outDir, true);
        Directory.CreateDirectory(outDir);
        ZipFile.ExtractToDirectory(inFile, outDir);
        Console.WriteLine($"Extracted {inFile} to {outDir}");
    }

    static void UnpackAuto(string inFile)
    {
        // Create folder named after the input file (without extension) in the same directory
        string directory = Path.GetDirectoryName(inFile) ?? throw new InvalidOperationException("Cannot determine directory");
        string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(inFile);
        string outDir = Path.Combine(directory, fileNameWithoutExtension);
        
        if (Directory.Exists(outDir)) Directory.Delete(outDir, true);
        Directory.CreateDirectory(outDir);
        ZipFile.ExtractToDirectory(inFile, outDir);
        Console.WriteLine($"Extracted {inFile} to {outDir}");
    }
}
