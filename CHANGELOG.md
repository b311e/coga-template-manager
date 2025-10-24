# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- Complete template management system for COGA agencies
- JBC theme standardization with proper color scheme
- XLSTART-compatible Excel templates (Book.xltx and Sheet.xltx)
- OpenXML document pack/unpack functionality
- Automated deployment scripts for PreProd and Production
- Bash utility scripts with alias support

### Features
- **Template Processing**: Pack/unpack .dotm and .xltx files
- **Theme Integration**: Standardized JBC colors and fonts
- **XLSTART Support**: Proper Excel default template structure
- **Multi-Agency Support**: Extensible framework for HOU, SEN, JBC, LCS, OLLS, OSA
- **Clean Deployment**: Automated folder cleanup and recreation

### Technical
- .NET 8.0 console application
- DocumentFormat.OpenXml 3.3.0 for template manipulation
- Windows batch scripts for deployment automation
- Bash scripts for development workflow

### Documentation
- Comprehensive README with setup instructions
- Contributing guidelines for developers
- Project structure documentation
- Troubleshooting guides

## Template Structure Changes

### JBC Templates
- **Book.xltx**: Full workbook template with JBC theme selected by default
- **Sheet.xltx**: Worksheet template for XLSTART compatibility
- **Theme Integration**: Proper x15ac:absPath and extLst elements for Excel recognition

### Deployment Structure
```
deploy/
├── core/                # Core deployment scripts
│   ├── PushToProd.bat
│   ├── PreProd_TemplateInstall.bat
│   └── TemplateInstall.bat
└── jbc/                 # Agency-specific
    └── JBCTemplateInstall.bat
```

### Development Tools
- Alias system for common commands
- XPath search functionality for XML debugging
- Automated template validation

---

## Future Considerations

### Planned Enhancements
- Additional agency template support
- Automated testing framework
- Template validation tools
- Enhanced error handling

### Breaking Changes
- None currently planned
- Will follow semantic versioning for any breaking changes

---

*This changelog will be updated as new versions are released.*