# Maintenance

Good to know:

- Leave lines 1-9 alone when editing logic. This is the Batch wrapper, and it's identical in both the pack and unpack scripts. If you need to change the wrapper, change it in **both** files to keep them in sync.

- Always keep `[System.IO.Path]` / `ZipArchive usage`, not "compress folder".  If you "simplifies" `pack` by calling `Compress-Archive` or `ZipFile.CreateFromDirectory`, it will introduce the backslash/extra-folder bugs and Office will reject the output (i.e., it will create an invalid file that won't open).

- `pack` decides the output **extension** by matching the main part's content type inside `[Content_Types].xml` (the `switch -Regex` block). To support a new format, add one `'<contenttype-fragment>' { '<ext>'; break }` line to that switch. Each fragment must uniquely identify the type; keep the `; break` so only the first match wins. 

    - **Order matters when adding new Office formats!** When adding an Office file format, put the most specific/unique marker earlier than any broader one it might also contain. Why? A `.thmx` theme, contains `presentationml.presentation.main` (it bundles presentation variants), so it would wrongly match `pptx` unless the `themeManager` line is placed **before** the `.pptx` line. 

- PowerShell note: inside a .NET method call, a comma separates one argument from the next. So if an argument itself contains a comma like `-replace 'a','b'` or `-split ','` — PowerShell reads it as several arguments and the call fails. Put extra parentheses around that argument so it stays as one: `GetFileNameWithoutExtension(($x -replace 'a',''))`. Plain arguments without a comma don't need this.