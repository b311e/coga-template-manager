# JBC Template Management Structure

## Overview
Organized template development structure for JBC (Johnson, Barrow & Company) templates and Office integration.

## Structure

### workspace/
Core Office integration templates that users interact with daily:
- **jbcNormal**: Word Normal.dotm template for default document creation
- **jbcBook**: Excel Book.xltx template for XLSTART folder
- **jbcSheet**: Excel Sheet.xltx template for new worksheet creation
- **jbcTheme**: Office theme files (.thmx) for consistent branding

### templates/
Document templates for specific business purposes:
- **jbcLetterhead**: Official letterhead template (.dotx)
- **jbcMemo**: Internal memo template (.dotx)
- **jbcForm**: Business forms and structured documents
- **jbcReport**: Report templates with JBC formatting

### system/
Agency-specific automation and deployment:
- Deployment scripts for JBC Office Workspace
- Theme installation automation  
- Custom template generators
- Agency-specific workflows

## Naming Convention

### Templates
All templates follow the pattern: `jbc{TemplateName}`
- Immediately identifies agency ownership
- Prevents conflicts with other agencies
- Maintains consistency across all templates

### Executables/Scripts
All scripts and executables follow the pattern: `{agency}{Target}{Action}`
- **Agency First**: `jbc` - immediately identifies ownership
- **Target Second**: `Workspace`, `Template`, `Theme` - describes what's being acted upon
- **Action Last**: `Deploy`, `Install`, `Setup` - describes the operation (imperative verb form)

## Workflow
Each template follows the src/in/out/docs pattern:
1. **src/**: Original template files
2. **in/**: Unpacked OpenXML for editing
3. **out/**: Generated files ready for deployment
4. **docs/**: Documentation and change tracking

## Scalability
This structure easily scales to multiple agencies:
```
builds/
├── jbc/
│   ├── workspace/
│   ├── templates/
│   └── system/
├── abc/
│   ├── workspace/
│   ├── templates/
│   └── system/
└── xyz/
    ├── workspace/
    ├── templates/
    └── system/
```

## Benefits
- **Clear organization** by function and purpose
- **Scalable** to multiple agencies
- **Consistent** naming and structure
- **Traceable** development with docs folders
- **Automated** with pack/unpack scripts