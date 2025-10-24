# Foundation Templates

The foundation directory contains the base templates and reusable components used to create agency-specific templates.

## Directory Structure

```
foundation/
├── base/                   # Complete base templates
│   ├── excel/             # Excel base templates
│   │   ├── Book-Base.xltx # Base workbook template
│   │   └── Sheet-Base.xltx # Base worksheet template
│   ├── word/              # Word base templates
│   │   └── Normal-Base.dotm # Base document template
│   └── themes/            # Base theme files
│       └── COGA-Base.thmx
└── partials/              # Reusable components
    ├── styles/            # Standard style definitions
    ├── xml-snippets/      # Common XML elements
    └── color-schemes/     # Color variations
```

## Usage

### Creating New Agency Templates

1. **Start with Base Template**
   ```bash
   cp foundation/base/excel/Book-Base.xltx templates/[agency]/book/template.xltx
   ```

2. **Customize for Agency**
   ```bash
   unpack templates/[agency]/book/template.xltx
   # Edit template/ folder with agency-specific changes
   pack template/ templates/[agency]/book/template.xltx
   ```

### Working on Foundation Templates

1. **Improve Base Templates**
   ```bash
   cd foundation/base/excel/
   unpack Book-Base.xltx
   # Make improvements to Book-Base/ folder
   pack Book-Base/ Book-Base.xltx
   ```

2. **Use Partials**
   ```bash
   # Copy standard components
   cp foundation/partials/styles/excel-styles.xml working-folder/xl/styles.xml
   ```

## Foundation Principles

- **Base templates** are complete, working templates ready for customization
- **Partials** are reusable components that can be mixed and matched
- Foundation templates use neutral styling that's easy to customize
- All templates maintain proper OpenXML structure and Office compatibility

---

*These are the foundational building blocks for all agency templates. Handle with care and test thoroughly.*