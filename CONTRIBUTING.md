# Contributing to COGA Template Manager

## Development Environment Setup

### Prerequisites
1. **Install .NET 8.0 SDK**
   - Download from [Microsoft .NET Downloads](https://dotnet.microsoft.com/download/dotnet/8.0)
   - Verify installation: `dotnet --version`

2. **Development Tools** (recommended)
   - Visual Studio Code with C# extension
   - Git for version control
   - Windows environment for testing deployment scripts

### Getting Started

1. **Clone and Setup**
   ```bash
   git clone https://github.com/b311e/coga-template-manager.git
   cd coga-template-manager
   dotnet restore
   dotnet build
   ```

2. **Set up Development Aliases**
   ```bash
   source src/scripts/setup_aliases.sh
   ```

## Project Architecture

### Core Components
- **OpenXmlApp**: .NET console application for template manipulation
- **Templates**: Source templates with expandable/packable structure
- **Deployment**: Batch scripts for PreProd/Prod deployment
- **Resources**: Reference files and defaults

### Template Structure
Templates follow this pattern:
```
templates/agency/templateName/
├── source/expanded/           # Unpacked template files
│   ├── [Content_Types].xml
│   ├── _rels/
│   ├── docProps/
│   └── xl/ (or word/)
└── template.xltx             # Packed template file
```

## Development Workflow

### Making Template Changes
1. **Unpack Template**
   ```bash
   unpack templates/jbc/jbcBook/template.xltx temp_folder
   ```

2. **Edit XML Files**
   - Modify `xl/workbook.xml`, `xl/styles.xml`, `xl/theme/theme1.xml`
   - Follow OpenXML standards
   - Maintain theme consistency

3. **Test Changes**
   ```bash
   pack temp_folder test_template.xltx
   # Test opening in Excel/Word
   ```

4. **Update Source**
   ```bash
   # Copy changes back to source
   cp -r temp_folder/* templates/jbc/jbcBook/source/expanded/
   pack templates/jbc/jbcBook/source/expanded templates/jbc/jbcBook/template.xltx
   ```

### Code Standards
- **XML**: Well-formed, properly namespaced
- **C#**: Follow .NET conventions, use meaningful variable names
- **Batch**: Include error handling, clear comments
- **Bash**: POSIX compliant where possible

### Testing
1. **Template Validation**
   - Ensure templates open in Office applications
   - Verify theme application
   - Test XLSTART compatibility

2. **Deployment Testing**
   - Test PreProd deployment scripts
   - Verify file copying and permissions
   - Test end-user installation

## Submitting Changes

### Pull Request Process
1. **Create Feature Branch**
   ```bash
   git checkout -b feature/description-of-change
   ```

2. **Make Changes**
   - Follow coding standards
   - Test thoroughly
   - Update documentation if needed

3. **Commit Guidelines**
   ```bash
   git commit -m "type: brief description
   
   More detailed explanation if needed.
   
   - Specific change 1
   - Specific change 2"
   ```

   **Commit Types:**
   - `feat:` New features
   - `fix:` Bug fixes
   - `docs:` Documentation changes
   - `style:` Formatting changes
   - `refactor:` Code restructuring
   - `test:` Test additions/changes
   - `deploy:` Deployment script changes

4. **Submit PR**
   - Clear title and description
   - Reference any related issues
   - Include testing notes

### Code Review
- All changes require review
- Focus on template compatibility
- Verify deployment script changes
- Test in representative environment

## Troubleshooting Development Issues

### Common Problems
- **OpenXML Errors**: Check XML syntax and namespaces
- **Theme Issues**: Verify styles.xml theme references
- **Build Failures**: Ensure .NET 8.0 SDK installed
- **Path Issues**: Use absolute paths in scripts

### Getting Help
- Check existing documentation in `docs/`
- Review similar implementations in codebase
- Contact team for architecture questions

## Agency-Specific Guidelines

When adding new agencies:
1. Create agency folder structure in `templates/`
2. Copy and modify deployment scripts
3. Update `PushToProd.bat` agency mappings
4. Test complete deployment workflow
5. Document agency-specific requirements