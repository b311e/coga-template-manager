# OOXML Content Types Cheat Sheet

The file type of an OOXML package is determined by the `ContentType` on the main-part `<Override>` in `[Content_Types].xml` at the root of the ZIP.

---

## Word (WordprocessingML)

| Extension | ContentType on `/word/document.xml` |
|-----------|-------------------------------------|
| `.docx`   | `application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml` |
| `.dotx`   | `application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml` |
| `.docm`   | `application/vnd.ms-word.document.macroEnabled.main+xml` |
| `.dotm`   | `application/vnd.ms-word.template.macroEnabledTemplate.main+xml` |

## Excel (SpreadsheetML)

| Extension | ContentType on `/xl/workbook.xml` |
|-----------|-----------------------------------|
| `.xlsx`   | `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml` |
| `.xltx`   | `application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml` |
| `.xlsm`   | `application/vnd.ms-excel.sheet.macroEnabled.main+xml` |
| `.xltm`   | `application/vnd.ms-excel.template.macroEnabled.main+xml` |
| `.xlam`   | `application/vnd.ms-excel.addin.macroEnabled.main+xml` |
| `.xlsb`   | `application/vnd.ms-excel.sheet.binary.macroEnabled.main` |

## PowerPoint (PresentationML)

| Extension | ContentType on `/ppt/presentation.xml` |
|-----------|----------------------------------------|
| `.pptx`   | `application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml` |
| `.potx`   | `application/vnd.openxmlformats-officedocument.presentationml.template.main+xml` |
| `.ppsx`   | `application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml` |
| `.pptm`   | `application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml` |
| `.potm`   | `application/vnd.ms-powerpoint.template.macroEnabled.main+xml` |
| `.ppsm`   | `application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml` |

---

## How to Read It

```xml
<!-- [Content_Types].xml -->
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <!-- The Override for the main part tells you the file type -->
  <Override PartName="/word/document.xml"
            ContentType="application/vnd.ms-word.template.macroEnabledTemplate.main+xml"/>
</Types>
```

The pattern in the ContentType string:
- `openxmlformats-officedocument` → open standard (no macros)
- `ms-word` / `ms-excel` / `ms-powerpoint` → Microsoft extension (macros or binary)
- `…template…` → template (`.dotx`, `.xltx`, `.potx`, etc.)
- `…macroEnabled…` / `…macroEnabled…` → macro-enabled (`.docm`, `.xlsm`, `.pptm`, etc.)

---

## Quick Detection (Python)

```python
import zipfile
from xml.etree import ElementTree as ET

NS = "http://schemas.openxmlformats.org/package/2006/content-types"

CONTENT_TYPE_TO_EXT = {
    # Word
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml": ".docx",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml": ".dotx",
    "application/vnd.ms-word.document.macroEnabled.main+xml":                          ".docm",
    "application/vnd.ms-word.template.macroEnabledTemplate.main+xml":                 ".dotm",
    # Excel
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml":     ".xlsx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml":  ".xltx",
    "application/vnd.ms-excel.sheet.macroEnabled.main+xml":                           ".xlsm",
    "application/vnd.ms-excel.template.macroEnabled.main+xml":                        ".xltm",
    # PowerPoint
    "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml": ".pptx",
    "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml":     ".potx",
    "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml":    ".ppsx",
    "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml":                   ".pptm",
}

def detect_extension(path: str) -> str | None:
    """Detect the correct file extension from the OOXML content type."""
    with zipfile.ZipFile(path) as z:
        with z.open("[Content_Types].xml") as f:
            tree = ET.parse(f)
    for override in tree.findall(f"{{{NS}}}Override"):
        ct = override.get("ContentType", "")
        if ct in CONTENT_TYPE_TO_EXT:
            return CONTENT_TYPE_TO_EXT[ct]
    return None
```

---

## References

- ECMA-376, 5th Edition — Part 2 §13 (Open Packaging Conventions, Content Types)
- Schema: `resources/schemas/ooxml/opc/opc-contentTypes.xsd`
