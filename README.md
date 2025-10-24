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
├── foundation/             # Foundation templates and components
│   ├── base/               # Complete base templates
│   └── partials/           # Reusable components
├── src/
│   ├── OpenXmlApp/         # Core .NET application for template processing
│   └── scripts/            # Bash utility scripts
├── builds/                 # Agency template development
│   └── jbc/                # JBC work-in-progress templates
├── dist/                   # Production-ready deployments
│   └── jbc/                # JBC Office Workspace (deployment-ready)
├── deploy/                 # Deployment scripts and configurations
│   ├── core/               # Core deployment scripts
│   └── jbc/                # Agency-specific scripts
├── resources/              # Reference files and defaults
└── docs/                   # Additional documentation
```

## Commands

### Template Processing
- `dotnet run --project src/OpenXmlApp pack <expanded_folder> <output_file>` - Pack expanded template folder into .dotm/.xltx
- `dotnet run --project src/OpenXmlApp unpack <template_file> [output_folder]` - Extract template to folder

### With Aliases (after running setup_aliases.sh)
- `pack <expanded_folder> <output_file>` - Pack template. Will save packed file to the build folder
- `unpack <template_file> [output_folder]` - Unpack template
- `xpathsearch <xpath>` - Search XML content with XPath

## Template Features

### Theme Integration
- **Standardized Colors**: Custom agency color scheme
- **Font Standards**: Agency branded font family with proper theme integration

### Supported Template Types
- **Book.xltx**: Default Excel workbook template
- **Sheet.xltx**: Default Excel worksheet template  
- **Normal.dotm**: Word document template

## 🏗️ Foundation System

### Foundation Templates
- **`foundation/base/`**: Complete base templates ready for agency customization
- **`foundation/partials/`**: Reusable components (themes, styles, XML snippets)

### Creating New Agency Templates
```bash
# Start from foundation
cp foundation/base/excel/Book-Base.xltx templates/newagency/book/template.xltx

# Customize for agency
unpack templates/newagency/book/template.xltx
# Edit template/ folder with agency branding
pack template/ templates/newagency/book/template.xltx
```

## 🚀 Deployment

### Development Workflow
1. Develop foundation templates in `foundation/base/`
2. Create agency templates from foundation
3. Customize for agency branding and requirements
4. Pack templates: `pack [expanded_folder] [output_file]`
5. Test templates locally
4. Deploy to PreProd for testing
5. Push to Production

### Deployment Scripts
- `deploy/core/PreProd_TemplateInstall.bat` - Deploy to PreProd environment
- `deploy/core/PushToProd.bat` - Deploy from PreProd to Production
- `deploy/jbc/JBCTemplateInstall.bat` - End-user installation script
- `style list <templateName>` – regenerate the style list for template. Saves to templateName/docs folder

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

### Creating a New Template
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
cd deploy\core
PreProd_TemplateInstall.bat

# After testing, push to Production
PushToProd.bat
```

## Contributing

1. Fork the repository
2. Create a feature branch: `git checkout -b feature/new-feature`
3. Make changes and test thoroughly
4. Commit with clear messages: `git commit -m "Add new feature"`
5. Push and create a Pull Request

## License

Internal Colorado General Assembly tool. Not for external distribution.

## Troubleshooting

### Common Issues
- **Build Failures**: Ensure .NET 8.0 SDK is installed
- **Theme Not Applied**: Verify styles.xml contains complete theme references

### Getting Help
- Check the `docs/` folder for additional documentation
- Review template XML structure in `resources/defaults/`
- Contact the development team for support