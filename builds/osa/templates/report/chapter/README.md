# Change Log

## 2025-12-18

osaReport is done. Re-build this template using that template tomorrow.

## 2025-12-02

Where I left off on the OSA Report Chapter Template:
- Deleted unused numbering styles
- Removed leader dots from TOC
- Compiled final
- Made edits to final during templates check in with OSA

--> Need to make requested changes (see below)

Deleted Numbering Styles:

```xml
    <w:style w:type="numbering"
             w:customStyle="1"
             w:styleId="Delete">
        <w:name w:val="Delete" />
        <w:uiPriority w:val="99" />
        <w:rsid w:val="000E3EE0" />
        <w:pPr>
            <w:numPr>
                <w:numId w:val="6" />
            </w:numPr>
        </w:pPr>
    </w:style>


    <w:style w:type="numbering"
             w:customStyle="1"
             w:styleId="Delete3">
        <w:name w:val="Delete3" />
        <w:uiPriority w:val="99" />
        <w:rsid w:val="00B446B1" />
        <w:pPr>
            <w:numPr>
                <w:numId w:val="2" />
            </w:numPr>
        </w:pPr>
    </w:style>


    <w:style w:type="numbering"
             w:customStyle="1"
             w:styleId="Delete2">
        <w:name w:val="Delete2" />
        <w:uiPriority w:val="99" />
        <w:rsid w:val="009D3432" />
        <w:pPr>
            <w:numPr>
                <w:numId w:val="4" />
            </w:numPr>
        </w:pPr>
    </w:style>
```

Additional changes were made to the out file during the meeting and need to be reapplied to the expanded in file, then recompiled.

**Requested updates (12/02 OSA Templates Check in):**

- Create new styles with the following style names and export tags:
	- First Subhead (H3)
	- Second Subhead (H4)
	- Third Subhead (H5)
- Change "Chapter #" to "Chapter Title"
- Change "Heading 2 Subtitle" to "Chapter Subtitle"
- Update other Subtitle style names as necessary

Create two Report Chapter Templates:
- Chapter 1 Template
	- Includes Audit, Purpose, Scope
	- Hide Recommendation Header
- Chapter Template
