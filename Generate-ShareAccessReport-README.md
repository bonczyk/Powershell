# Generate-ShareAccessReport

A professional PowerShell function for generating enterprise-grade share access reports in multiple formats (HTML, XLSX, PDF).

## Overview

`Generate-ShareAccessReport` is a comprehensive reporting tool designed for corporate and banking environments. It processes share access data and produces professional, formatted reports with:

- **Multiple output formats**: HTML, Excel (XLSX), and PDF
- **Flexible grouping**: By Owner or by Server
- **Interactive features**: Expandable/collapsible sections
- **Professional themes**: Three built-in corporate themes
- **Rich analytics**: Summary statistics and risk assessment
- **Enterprise styling**: Clean, formal designs suitable for executive presentation

## Requirements

### PowerShell Version
- PowerShell 5.1 or higher

### Required Modules
```powershell
# Install ImportExcel module for Excel generation
Install-Module -Name ImportExcel -Scope CurrentUser -Force

# Install PSWritePDF module for PDF generation
Install-Module -Name PSWritePDF -Scope CurrentUser -Force
```

## Features

### Report Types

#### PerOwner Mode
Groups data by share owners (Owner1 and Owner2), showing:
- All shares owned by each person
- AD groups with access to those shares
- Nested lists of users in each group
- Complete access rights breakdown

#### PerServer Mode
Groups data by server, showing:
- All shares on each server
- Share ownership information
- AD groups with access
- User details and rights

### Themes

Three professional themes are available:

1. **CorporateBlue** (Default)
   - Modern corporate design with blue gradients
   - Professional sans-serif fonts
   - Suitable for financial institutions

2. **MinimalGray**
   - Clean, minimalist design with gray tones
   - Helvetica-based typography
   - Ideal for technical documentation

3. **ExecutiveGreen**
   - Executive-level presentation style
   - Serif fonts with green accents
   - Perfect for management reports

### Output Features

#### HTML Reports
- Responsive, professional design
- Color-coded risk levels (High/Medium/Low)
- Interactive expandable/collapsible sections
- Executive summary with key metrics
- Professional headers and footers

#### Excel Reports
- Multiple worksheets for organized data
- Summary worksheet with key statistics
- Full data export
- Separate worksheets per server/owner
- Auto-sized columns
- Bold headers with freeze panes
- Conditional formatting for high-risk permissions
- Professional table styling

#### PDF Reports
- High-quality rendering
- Professional page layout
- Proper margins and spacing
- Headers and footers with report metadata
- Suitable for printing and distribution

## Usage

### Basic Syntax

```powershell
Generate-ShareAccessReport -Data <PSCustomObject[]> -ReportType <String> [-Theme <String>] 
    [-Expandable] [-HtmlPath <String>] [-XlsxPath <String>] [-PdfPath <String>]
```

### Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| Data | PSCustomObject[] | Yes | Array of share access data objects |
| ReportType | String | Yes | "PerOwner" or "PerServer" |
| Theme | String | No | "CorporateBlue", "MinimalGray", or "ExecutiveGreen" (Default: CorporateBlue) |
| Expandable | Switch | No | Makes report sections collapsible |
| HtmlPath | String | No | Output path for HTML report (Default: timestamped in current directory) |
| XlsxPath | String | No | Output path for Excel report (Default: timestamped in current directory) |
| PdfPath | String | No | Output path for PDF report (Default: timestamped in current directory) |

### Data Structure

The input data must be an array of PSCustomObjects with the following properties:

```powershell
[PSCustomObject]@{
    Server      = "ServerName"          # File server name
    Share       = "ShareName"           # Share name
    ADGroupName = "GroupName"           # Active Directory group name
    Domain      = "DomainName"          # User domain
    User        = "username"            # User account name
    DisplayName = "Full Name"           # User display name
    UserGroup   = "DepartmentName"      # User's group/department
    SharePath   = "D:\Path\To\Share"    # Physical path
    Owner1      = "Primary Owner"       # Primary share owner
    Owner2      = "Secondary Owner"     # Secondary owner (optional)
    Rights      = "FullControl"         # Access rights
}
```

## Examples

### Example 1: Basic PerOwner Report

```powershell
# Import the function
. .\Generate-ShareAccessReport.ps1

# Generate report grouped by owner
$result = Generate-ShareAccessReport `
    -Data $ExpandedData `
    -ReportType "PerOwner" `
    -Theme "CorporateBlue"

# View summary
$result.Summary
```

### Example 2: PerServer Report with All Formats

```powershell
# Generate comprehensive report with custom paths
Generate-ShareAccessReport `
    -Data $ExpandedData `
    -ReportType "PerServer" `
    -Theme "ExecutiveGreen" `
    -HtmlPath "C:\Reports\ShareAccess_Server.html" `
    -XlsxPath "C:\Reports\ShareAccess_Server.xlsx" `
    -PdfPath "C:\Reports\ShareAccess_Server.pdf"
```

### Example 3: Expandable Report with MinimalGray Theme

```powershell
# Create collapsible report sections
Generate-ShareAccessReport `
    -Data $ExpandedData `
    -ReportType "PerOwner" `
    -Theme "MinimalGray" `
    -Expandable `
    -Verbose
```

### Example 4: Integration with Existing ACL Scan

```powershell
# After running your ACL scan (similar to Get-ACLScan2.ps1)
$ScriptPath = $PSScriptRoot
$ReportsPath = "$ScriptPath\Reports"
$Timestamp = Get-Date -Format "yyyyMMdd_HHmm"

# ... your ACL scanning and group expansion code ...
# Result in $ExpandedData array

# Generate both report types
Generate-ShareAccessReport `
    -Data $ExpandedData `
    -ReportType "PerServer" `
    -Theme "CorporateBlue" `
    -Expandable `
    -HtmlPath "$ReportsPath\Server_Report_$Timestamp.html" `
    -XlsxPath "$ReportsPath\Server_Report_$Timestamp.xlsx" `
    -PdfPath "$ReportsPath\Server_Report_$Timestamp.pdf"

Generate-ShareAccessReport `
    -Data $ExpandedData `
    -ReportType "PerOwner" `
    -Theme "ExecutiveGreen" `
    -HtmlPath "$ReportsPath\Owner_Report_$Timestamp.html" `
    -XlsxPath "$ReportsPath\Owner_Report_$Timestamp.xlsx" `
    -PdfPath "$ReportsPath\Owner_Report_$Timestamp.pdf"

# Open reports folder
Invoke-Item $ReportsPath
```

### Example 5: Empty Data Handling

```powershell
# Function handles empty data gracefully
Generate-ShareAccessReport `
    -Data @() `
    -ReportType "PerServer" `
    -Theme "CorporateBlue" `
    -HtmlPath "C:\Reports\EmptyReport.html"
```

## Report Contents

### Executive Summary Section

Every report includes:
- Total number of records
- Unique servers count
- Unique shares count
- Unique owners count
- Unique users count
- Unique AD groups count
- High-risk permissions count

### Data Hierarchy

#### PerOwner Reports
```
Owner
└── Share
    └── AD Group
        └── Users
            - Domain
            - Username
            - Display Name
            - User Group
            - Rights
```

#### PerServer Reports
```
Server
└── Share (with Owner info)
    └── AD Group
        └── Users
            - Domain
            - Username
            - Display Name
            - User Group
            - Rights
```

### Risk Indicators

Permissions are automatically color-coded:
- **High Risk** (Red): FullControl permissions
- **Medium Risk** (Orange): Modify, Write permissions
- **Low Risk** (Green): Read, Execute permissions

## Output Files

### HTML Reports
- Self-contained with embedded CSS
- Viewable in any modern web browser
- Interactive elements (expandable sections)
- Print-friendly layout
- Professional branding placeholders

### Excel Reports
- Multiple worksheets for easy navigation
- Summary tab with key metrics
- Full data export on separate sheet
- Individual sheets per server/owner (with name truncation for Excel limits)
- Professional table styling with alternating row colors
- Conditional formatting highlighting high-risk permissions
- Auto-sized columns for readability
- Frozen header rows

### PDF Reports
- Converted from HTML with professional layout
- Suitable for distribution and archival
- Includes all visual formatting from HTML
- Proper page breaks
- Headers and footers with report metadata

## Best Practices

1. **Module Installation**: Install required modules before first use
2. **Large Datasets**: For very large datasets (>10,000 records), consider filtering data or generating separate reports
3. **Path Management**: Use dedicated reports folders with timestamps for version control
4. **Theme Selection**: Choose theme based on audience (Executive for management, Corporate for general use, Minimal for technical teams)
5. **Regular Execution**: Schedule regular report generation for compliance and audit purposes
6. **Expandable Option**: Use `-Expandable` for very large reports to improve initial load time

## Troubleshooting

### Module Import Errors
```powershell
# Verify modules are installed
Get-Module -ListAvailable ImportExcel, PSWritePDF

# Install if missing
Install-Module -Name ImportExcel -Scope CurrentUser -Force
Install-Module -Name PSWritePDF -Scope CurrentUser -Force
```

### PDF Generation Issues
If PDF generation fails, the HTML report can be manually converted:
1. Open the HTML report in a web browser
2. Use browser's "Print to PDF" function
3. Or use a third-party HTML-to-PDF converter

### Large File Performance
For very large datasets:
- Use `-Expandable` switch to collapse sections by default
- Consider splitting data by server or owner before reporting
- Generate HTML only first, then Excel/PDF separately if needed

### Excel Worksheet Name Errors
Excel worksheet names are automatically sanitized:
- Invalid characters are replaced with underscores
- Names are truncated to 31 characters (Excel limit)
- This is handled automatically by the function

## Security Considerations

- Reports may contain sensitive access information
- Store reports in secure locations
- Implement appropriate access controls
- Consider encryption for sensitive environments
- Add custom confidentiality warnings as needed
- Review data before distribution to external parties

## Integration with Existing Scripts

This function is designed to integrate with existing ACL scanning scripts in the repository:

```powershell
# After your existing ACL scan (Get-ACLScan2.ps1 pattern)
# You have $ExpandedData populated

# Generate reports
. "$PSScriptRoot\Generate-ShareAccessReport.ps1"

Generate-ShareAccessReport `
    -Data $ExpandedData `
    -ReportType "PerServer" `
    -Theme "CorporateBlue" `
    -HtmlPath "$ReportsPath\ShareReport_$Timestamp.html" `
    -XlsxPath "$ReportsPath\ShareReport_$Timestamp.xlsx"
```

## Customization

### Adding Custom Themes
Edit the `Get-ThemeStyles` function to add new themes by copying an existing theme and modifying colors, fonts, and spacing.

### Modifying Report Structure
The HTML generation functions (`Generate-PerOwnerHtml`, `Generate-PerServerHtml`) can be customized to change the report structure and hierarchy.

### Custom Summary Metrics
Modify the `Get-ReportSummary` function to include additional statistics relevant to your organization.

## Version History

- **Version 1.0** (Initial Release)
  - PerOwner and PerServer report types
  - Three professional themes
  - HTML, XLSX, and PDF output
  - Expandable/collapsible sections
  - Executive summary
  - Risk-based color coding
  - Empty data handling

## License

This function is provided as-is for use in enterprise environments. Modify as needed for your specific requirements.

## Support

For issues, enhancements, or questions:
1. Review this documentation
2. Check the example script for usage patterns
3. Review function comments and verbose output
4. Test with sample data first

## Related Scripts

- `Get-ACLScan` - Original ACL scanning script
- `Get-ACLScan2` - Enhanced ACL scanning with AD integration
- Both can provide data in the format expected by this function

## Author

Enterprise PowerShell Team - Specialized in creating professional reporting tools for corporate environments.
