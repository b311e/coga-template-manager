# templx-lite

**Version:** 2026-06-11_v1.0.0

A Word/Excel/PowerPoint file (`.docx`, `.dotm`, `.xlsx`, `.xltx`, etc.) or `.thmx` file is really just a `.zip` of XML. These scripts help you open and edit them.

## Usage

### 1. `unpack.bat`: unzip the file

**Drag and drop** a Word/Excel/PowerPoint file onto **`unpack.bat`**.

It extracts the file into a sibling folder next to it:

```
Report.docx    ->    Report.docx_unpacked\
```

### 2. `pack.bat`: put it back together

**Drag and drop** the **`_unpacked` folder** onto **`pack.bat`**.

It rebuilds a valid Office file next to the unpacked folder:

```
Report.docx_unpacked\  ->  Report_packed.docx
```