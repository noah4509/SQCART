# SQCART - Supplier Quality Corrective Action Tool

![SQCART Logo](docs/images/sqcart_logo.png)

## Overview
SQCART is a tool designed to analyze supplier corrective action responses using the 8D (Eight Disciplines) Problem-Solving Process against the AS9100 standard. It provides automated feedback on gaps and suggests improvements.

## Quick Installation
1. Download `SQCART_Install.bat`
2. Right-click the downloaded file and select "Run as administrator"
3. Follow the on-screen prompts

![Installation Process](docs/images/install_process.png)

## Requirements
- Windows 7/10/11
- Microsoft Excel 2016 or newer
- Administrative rights for installation
- Minimum 100MB free disk space
- .NET Framework 4.5 or higher

## Technical Details

### System Integration
- Integrates with Microsoft Excel using COM automation
- Creates custom UI elements in Excel's Quick Access Toolbar
- Utilizes Windows PowerShell for installation tasks
- Implements VBA for core functionality

### File Structure
```text
%PROGRAMFILES%\SQCART\
├── SQCART_Template.xlsm    # Main Excel template
├── uninstall.bat           # Uninstallation script
└── config/                 # Configuration files

%USERPROFILE%\Documents\SQCART\
├── Reports/               # PDF exports
└── Templates/            # Custom templates
```

### Excel Template Architecture
- **Supplier Input Form**: Data entry and validation
- **AS9100 Compliance Check**: Analysis engine
- **Report**: Automated report generation

![Excel Template Structure](docs/images/template_structure.png)

## Features
### 8D Analysis
- Cross-functional team formation tracking
- Problem description validation
- Root cause analysis tools
- Corrective action planning
- Implementation verification
- Preventive measures assessment

### PDF Export Functionality
1. Quick Access Toolbar Integration
   ![PDF Export Button](docs/images/pdf_export_button.png)
2. Automated File Naming
   - Format: `SQCART_Report_YYYY-MM-DD_HHMMSS.pdf`
   - Customizable naming conventions
3. Export Options
   - Quality settings: Standard/High
   - Include/exclude document properties
   - Custom page ranges

![PDF Export Dialog](docs/images/pdf_export_dialog.png)

## Troubleshooting Guide

### Installation Issues

#### Error: "Access Denied"
**Symptoms:**
- Installation fails with access denied message
- Unable to create program directories

**Solutions:**
1. Right-click installer and select "Run as administrator"
2. Temporarily disable antivirus
3. Check Windows UAC settings

#### Error: "Excel Not Found"
**Symptoms:**
- Installation warns about missing Excel
- Template fails to open

**Solutions:**
1. Install Microsoft Excel 2016 or newer
2. Repair Office installation
3. Run Office updates

![Excel Error Solution](docs/images/excel_error_fix.png)

### PDF Export Issues

#### Error: "Unable to Create PDF"
**Symptoms:**
- PDF export fails
- No error message appears

**Solutions:**
1. Check folder permissions
2. Close existing PDF viewers
3. Clear temporary files
```powershell
# Run in PowerShell as administrator
Remove-Item "$env:TEMP\*.tmp" -Force
```

#### Error: "VBA Project Error"
**Symptoms:**
- Macro security warnings
- PDF button missing

**Solutions:**
1. Enable macros:
   - File → Options → Trust Center → Trust Center Settings
   - Enable "Trust access to the VBA project object model"
   ![Macro Settings](docs/images/macro_settings.png)

2. Reset Excel settings:
```batch
excel.exe /safe
```

### Common Issues and Solutions

| Issue | Cause | Solution |
|-------|-------|----------|
| Slow PDF Export | Large file size | Reduce sheet range |
| Missing Toolbar | Excel customization | Reset UI elements |
| File Lock Error | File in use | Close other instances |

## Performance Optimization
- Clear unused sheets
- Limit data ranges
- Regular temp file cleanup
- Update Excel regularly

![Performance Tips](docs/images/performance_tips.png)

## Support
If you encounter any issues:
1. Check the troubleshooting guide above
2. Review system requirements
3. Try running in safe mode:
   ```batch
   "%PROGRAMFILES%\SQCART\SQCART_Template.xlsm" /safe
   ```
4. Open an issue on GitHub with:
   - Windows version
   - Excel version
   - Error screenshots
   - Steps to reproduce

## Version History
- 1.0.0 - Initial Release
  - Basic 8D analysis
  - PDF export feature
- 1.0.1 - Bug fixes
  - Fixed PDF naming issue
  - Improved error handling

## License
MIT License - See LICENSE file for details

## Contributing
Feel free to submit issues and pull requests for improvements.