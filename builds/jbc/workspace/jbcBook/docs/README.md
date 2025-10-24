# Book Template Documentation

## Overview
JBC Excel workbook template for default Book.xltx in XLSTART folder.

## Template Details
- **Type**: Excel Workbook Template (.xltx)
- **Purpose**: Default template when users create new Excel workbooks
- **Deployment**: XLSTART folder as Book.xltx
- **Theme**: JBC Office theme with standardized colors and fonts

## Development Notes
- Source: `src/Book.xltx` - Original template file
- Working: `in/` - Expanded XML files for editing
- Output: `out/Book.xltx` - Built template ready for testing
- Documentation: `docs/` - This folder

## Key Features
- JBC theme automatically selected
- Proper XLSTART compatibility with x15ac:absPath and extLst elements
- Complete Excel styles with theme color references
- Window dimensions optimized for standard displays

## Testing Checklist
- [ ] Template opens in Excel without errors
- [ ] JBC theme shows as selected in Design tab
- [ ] Colors and fonts apply correctly
- [ ] Cell styles available in Home tab
- [ ] Compatible with different Excel versions

## Deployment Path
`builds/jbc/book/out/Book.xltx` → `dist/jbc/JBC Office Workspace/JBC XLSTART/Book.xltx`

## Last Updated
Template last modified: [Date]
Documentation last updated: [Date]

## Change Log
- Initial JBC theme integration
- XLSTART compatibility improvements
- [Add your changes here]