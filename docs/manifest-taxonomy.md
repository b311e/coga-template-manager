# COGA Template Manager Manifest Taxonomy

## Version: 1.0

This document defines the complete taxonomy (classification system) for the COGA Template Manager manifest system.

## Core Structure

### Agency Level
- **agency**: Lowercase agency code (jbc, sen, hou, etc.)
- **agency_name**: Full agency name
- **manifest_version**: Semantic version (1.0, 1.1, etc.)
- **generated**: ISO 8601 timestamp

### Categories
Each agency manifest contains these top-level categories:

1. **workspace** - Office workspace templates (Book, Sheet, Normal)
2. **templates** - Document templates (Letterhead, Memo, Reports)
3. **system** - Deployment and automation scripts
4. **assets** - Themes, fonts, colors, deployment components

## Template Types

### Excel Templates
- `excel-book-template` - Excel workbook templates (.xltx)
- `excel-sheet-template` - Excel worksheet templates (.xltx)
- `excel-book` - Excel workbooks (.xlsx)

### Word Templates  
- `word-doc-template` - Word document templates (.dotx or .dotm)
- `word-doc` - Word documents (.docx)

## Status Values

Template development status:

- **active** - Template is complete and ready for use
- **planned** - Template is planned but not yet developed
- **testing** - Template exists but is in testing phase
- **deprecated** - Template is obsolete and should not be used

## File Structure

Each template entry includes these path references:

- **src** - Source template file (original)
- **out** - Built/processed template file
- **expanded** - Unpacked/expanded template directory
- **docs** - Documentation directory

Example:
```
builds/jbc/workspace/jbcBook/
├── src/Book.xltx          # Source template
├── out/Book.xltx          # Built template  
├── in/                    # Expanded template
│   ├── [Content_Types].xml
│   ├── _rels/
│   └── xl/
└── docs/                  # Documentation
    └── README.md
```

## Template Naming Convention

### Template IDs
- Format: `{agency}{TemplateType}`
- Examples: `jbcBook`, `jbcNormal`, `senLetterhead`
- Pattern: `^[a-zA-Z][a-zA-Z0-9]*$`

### File Extensions
- `.xltx` - Excel template
- `.dotx` - Word template (newer format)
- `.dotm` - Word template with macros
- `.xlsx` - Excel workbook
- `.docx` - Word document

## System Scripts

Scripts in the `system` category:

- **type**: Script type (deployment, automation, etc.)
- **file**: Path to script file
- **target**: Target audience (end-user-install, admin, etc.)
- **status**: Same status values as templates

## Assets Structure

Deployment assets for Office workspace:

- **workspace_dist** - Base distribution directory
- **components** - Individual component directories
  - Theme files, fonts, colors, Office UI customizations

## Validation

Manifests can be validated against the JSON schema:
`.coga/schemas/manifest-schema.json`

Use tools like `ajv` or online validators to check manifest structure.

## Evolution

### Adding New Template Types
1. Update schema: `.coga/schemas/manifest-schema.json`
2. Update script: `src/scripts/manifest` (add-template function)
3. Update documentation: This file and README.md

### Adding New Status Values
1. Update schema: `template_types` enum
2. Update script: `update_template_status` function
3. Update documentation

### Adding New Categories
1. Update global registry: `.coga/registry/agencies.json`
2. Update schema: `categories` enum  
3. Update agency manifests with new category structure

## Schema Location

The formal JSON schema is located at:
`.coga/schemas/manifest-schema.json`

This provides machine-readable validation rules for all manifest files.