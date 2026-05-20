using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

class Program
{
    static int Main(string[] args)
    {
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: style-import-doc <target-docx> <source-docx> [--dry-run] [--backup]");
            return 2;
        }

        var target = args[0];
        var source = args[1];
        var dryRun = Array.Exists(args, a => a == "--dry-run");
        var backup = Array.Exists(args, a => a == "--backup");

        if (!File.Exists(source))
        {
            Console.WriteLine("Source document not found: " + source);
            return 3;
        }

        if (!File.Exists(target))
        {
            Console.WriteLine("Target document not found: " + target);
            return 3;
        }

        if (dryRun)
        {
            Console.WriteLine($"Dry run: would replace styles in '{target}' from '{source}'");
            return 0;
        }

        string backupPath = target + ".bak";
        if (backup)
        {
            File.Copy(target, backupPath, overwrite: true);
        }

        try
        {
            using (var src = WordprocessingDocument.Open(source, false))
            using (var tgt = WordprocessingDocument.Open(target, true))
            {
                var srcStyles = src.MainDocumentPart?.StyleDefinitionsPart?.Styles;
                if (srcStyles == null)
                {
                    Console.WriteLine("Source document has no styles part: " + source);
                    return 4;
                }

                var tgtPart = tgt.MainDocumentPart?.StyleDefinitionsPart;
                if (tgtPart == null)
                {
                    // create styles part if missing
                    tgtPart = tgt.MainDocumentPart!.AddNewPart<StyleDefinitionsPart>();
                    tgtPart.Styles = new Styles();
                }

                // Replace target styles with a clone of the source styles
                tgtPart.Styles = (Styles)srcStyles.CloneNode(true);
                tgtPart.Styles.Save();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
            return 5;
        }

        if (backup)
            Console.WriteLine($"Wrote {target} (backup at {backupPath})");
        else
            Console.WriteLine($"Wrote {target}");

        return 0;
    }
}
