# COGA Template Manager

A comprehensive template management system for Colorado General Assembly Office templates, including Word (.dotm) and Excel (.xltx) templates with standardized branding and themes.

## Quick Start

### Prerequisites
- [.NET 8.0 SDK](https://dotnet.microsoft.com/download/dotnet/8.0)
- Windows environment (for deployment scripts)
- Access to CGA S Drive (for production deployment)

### Installation
1. Clone the repository:
   ```bash
   git clone https://github.com/b311e/coga-template-manager.git
   cd coga-template-manager
   ```

2. Build the application:
   ```bash
   dotnet build
   ```

3. Set up bash aliases (optional):
   ```bash
   source src/scripts/setup_aliases.sh
   ```

## Project Structure

```
coga-template-manager/
├── .coga/                  # COGA system metadata (hidden)
│   ├── registry/           # Agency registry and global configuration
│   └── schemas/            # Manifest validation schemas
├── foundation/             # Foundation templates and components
│   ├── base/               # Complete base templates
│   └── partials/           # Reusable components
├── src/
│   ├── OpenXmlApp/         # Core .NET application for template processing
│   └── scripts/            # Bash utility scripts (pack, unpack, create)
├── builds/                 # Agency template development
│   ├── jbc/                # JBC agency templates
│   │   ├── workspace/      # Office workspace templates (jbcBook, jbcNormal, jbcSheet)
│   │   ├── templates/      # Document templates (jbcLetterhead, jbcMemo)
│   │   └── system/         # Deployment and automation scripts
│   └── sen/                # Senate agency templates (similar structure)
├── dist/                   # Production-ready deployments
│   ├── jbc/
│   │   └── jbcWorkspace/   # JBC Office Workspace (jbcNormal, jbcFonts, etc.)
│   └── scripts/            # Core deployment scripts
├── resources/              # Reference files and documentation
└── docs/                   # Additional documentation
    └── manifest-taxonomy.md # Manifest system documentation
```

## Commands

### Template Processing
- `dotnet run --project src/OpenXmlApp pack <expanded_folder> <output_file>` - Pack expanded template folder into .dotm/.xltx
- `dotnet run --project src/OpenXmlApp unpack <template_file> [output_folder]` - Extract template to folder
- `dotnet run --project src/OpenXmlApp create <type> [name]` - Create new templates from scratch

### Template Registry
- `manifest list` - Show all templates across all agencies
- `manifest validate` - Validate all manifest files
- `manifest generate <agency>` - Generate/update manifest for specific agency
- `manifest update-status <agency> <template> <status>` - Update template status
- `manifest add-template <agency> <template> <type>` - Guide for adding new template

### With Aliases (after running setup_aliases.sh)
- `create <type> [name]` - Create new OpenXML templates/documents
  - Types: `excel-book-template`, `excel-sheet-template`, `excel-book`, `word-doc-template`, `word-doc`
- `manifest <command> [options]` - Manage template manifests and registry
- `pack <expanded_folder> <output_file>` - Pack template. Will save packed file to the build folder
- `unpack <template_file> [output_folder]` - Unpack template
- `style_list list <templateName>` - Generate style list for template

## Template Features

### Theme Integration
- **Standardized Colors**: Custom agency color scheme
- **Font Standards**: Agency branded font family with proper theme integration

### Supported Template Types
- **Book.xltx**: Default Excel workbook template
- **Sheet.xltx**: Default Excel worksheet template  
- **Normal.dotm**: Word document template

## Foundation System

### Foundation Templates
- **`foundation/base/`**: Complete base templates ready for agency customization
- **`foundation/partials/`**: Reusable components (themes, styles, XML snippets)

### Agency Structure
- **`builds/{agency}/workspace/`**: Office workspace templates (Book, Sheet, Normal)
- **`builds/{agency}/templates/`**: Document templates (Letterhead, Memo, etc.)
- **`builds/{agency}/system/`**: Deployment and automation scripts

### Creating New Agency Templates
```bash
# Create from scratch
create excel-book-template agency-book

# Or start from foundation
cp foundation/base/excel/Book-Base.xltx builds/newagency/workspace/agencyBook/src/Book.xltx

# Customize for agency
unpack builds/newagency/workspace/agencyBook/src/Book.xltx
# Edit expanded folder with agency branding
pack Book/ builds/newagency/workspace/agencyBook/out/Book.xltx
```

## Template Registry & Manifests

### Manifest Structure
The system uses JSON manifests to track templates, assets, and deployment configurations:

- **Global Registry**: `.coga/registry/agencies.json` - Lists all agencies and their manifest locations
- **Agency Manifests**: `builds/{agency}/manifest.json` - Complete template inventory per agency
- **Schema & Taxonomy**: `.coga/schemas/manifest-schema.json` - Validation schema, `docs/manifest-taxonomy.md` - Documentation
- **Category Structure**: Templates organized by workspace, templates, system, and assets

### Manifest Commands
```bash
# List all templates across all agencies
manifest list

# Validate all manifest files
manifest validate

# Generate/update manifest for specific agency
manifest generate jbc

# Update template status
manifest update-status jbc jbcNormal active

# Get guidance for adding new templates
manifest add-template jbc jbcReport excel-book-template
```

### Updating Manifests

#### Quick Updates
```bash
# Update timestamp and scan for new templates
manifest generate jbc

# Update template status (active, planned, deprecated, testing)
manifest update-status jbc jbcNormal active
```

## Template Inventory System

### Generate Template Inventory
The inventory system provides a user-friendly overview of all templates across all agencies:

```bash
# Generate template inventory report
coga inventory generate

# Or directly
inventory generate

# Show help
inventory help
```

### Inventory Output
The inventory generates `docs/template-inventory.md` with:
- **Current Status**: Real-time template status pulled from manifests
- **All Agencies**: Complete listing of House, Senate, JBC, LCS, OLLS, OSA
- **Template Details**: Template names and current status (active, planned, testing, deprecated)
- **Timestamp**: Generated date for tracking currency

### Example Output
```markdown
# Template Inventory & Status

Generated: 2025-10-24T07:56:48Z

## Joint Budget Committee

- jbcBook -- active
- jbcSheet -- active  
- jbcNormal -- active
- jbcLetterhead -- planned
- jbcMemo -- planned

## Senate

*No templates defined*
```

The inventory automatically reads from the manifest system to ensure accuracy and provides an at-a-glance view of template development progress across all agencies.

### Updating the Inventory List

To update the template inventory table in `docs/template-inventory.md` with the latest templates and statuses from all agencies, run:

```bash
coga inventory generate
```

This command will:
- Scan all agency manifests for current template names and statuses
- Regenerate the markdown table for each agency (including empty tables if no templates are defined)
- Overwrite the inventory file with the latest data and timestamp

You can also run the script directly:
```bash
src/scripts/inventory generate
```

**Note:** If you add, remove, or change templates or statuses in any agency manifest, rerun the inventory command to update the list.

For help:
```bash
coga inventory help
```

#### Adding New Templates
1. **Create template structure**:
   ```bash
   mkdir -p builds/jbc/workspace/jbcReport/{src,out,in,docs}
   ```

2. **Edit manifest manually**:
   ```bash
   code builds/jbc/manifest.json
   ```

3. **Add template entry** in appropriate section:
   ```json
   "jbcReport": {
     "name": "JBC Report Template",
     "type": "excel-book-template",
     "extension": ".xltx",
     "src": "builds/jbc/workspace/jbcReport/src/Report.xltx",
     "out": "builds/jbc/workspace/jbcReport/out/Report.xltx",
     "expanded": "builds/jbc/workspace/jbcReport/in/",
     "docs": "builds/jbc/workspace/jbcReport/docs/",
     "status": "planned"
   }
   ```

4. **Update manifest**:
   ```bash
   manifest generate jbc
   ```

#### Backup & Recovery
- Status updates create automatic backups: `manifest.json.backup`
- Manual backup: `cp builds/jbc/manifest.json builds/jbc/manifest.json.$(date +%Y%m%d)`

### Manifest Benefits
- **Template Discovery**: Easily find all templates and their locations
- **Status Tracking**: Track development status (active, planned, deprecated, testing)
- **Deployment Mapping**: Maps build templates to distribution locations
- **Documentation Links**: References to docs and specifications
- **Automated Updates**: Timestamp tracking and template discovery

## Deployment

### Development Workflow
1. Develop foundation templates in `foundation/base/`
2. Create agency templates from foundation
3. Customize for agency branding and requirements
4. Pack templates: `pack [expanded_folder] [output_file]`
5. Test templates locally
4. Deploy to PreProd for testing
5. Push to Production

### Deployment Scripts
- `dist/scripts/workspaceInstallPreProd.bat` - Deploy to PreProd environment
- `dist/scripts/workspaceDeploy.bat` - Deploy from PreProd to Production
- `dist/jbc/jbcWorkspace/JBCTemplateInstall.bat` - End-user installation script
- `style_list list <templateName>` – Generate style list for template. Saves to templateName/docs folder

## Development

### Dependencies
- **.NET 8.0**: Core runtime
- **DocumentFormat.OpenXml 3.3.0**: OpenXML document manipulation
- **Windows**: Required for batch deployment scripts

### Building from Source
```bash
# Restore dependencies
dotnet restore

# Build solution
dotnet build

# Run tests (if any)
dotnet test
```

## Usage Examples

### Creating New Templates from Scratch
```bash
# Create Excel book template
create excel-book-template mybook
# Creates: mybook.xltx

# Create Word document template  
create word-doc-template letter
# Creates: letter.dotx

# Create with default names
create excel-sheet-template
# Creates: Sheet.xltx
```

### Managing Template Manifests
```bash
# View all templates across agencies
manifest list

# Update template status
manifest update-status jbc jbcNormal active
manifest update-status jbc jbcBook testing

# Add new template (shows guidance)
manifest add-template jbc jbcReport excel-book-template

# Update manifest after changes
manifest generate jbc

# Validate all manifests
manifest validate

# Creates: Sheet.xltx
```

### Working with Existing Templates
```bash
# Unpack existing template (creates folder named after the template)
unpack existing_template.xltx
# This creates "existing_template/" folder

# Edit files in existing_template/
# Modify XML, themes, styles as needed

# Pack back to template
pack existing_template/ new_template.xltx
```

### Deploying JBC Templates
```cmd
# Deploy to PreProd for testing
cd dist\scripts
workspaceInstallPreProd.bat

# After testing, push to Production
workspaceDeploy.bat
```

## Contributing

1. Fork the repository
2. Create a feature branch: `git checkout -b feature/new-feature`
3. Make changes and test thoroughly
4. Commit with clear messages: `git commit -m "Add new feature"`
5. Push and create a Pull Request

### Common Issues
- **Build Failures**: Ensure .NET 8.0 SDK is installed
- **Theme Not Applied**: Verify styles.xml contains complete theme references

### Getting Help
- Check the `docs/` folder for additional documentation
- Review template XML structure in `resources/defaults/`
- Contact the development team for support