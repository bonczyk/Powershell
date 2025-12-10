# ShareAccessReport-Extensions

Extended functionality for Generate-ShareAccessReport to create per-owner reports and automated email confirmations for share access reviews.

## Overview

This module extends the `Generate-ShareAccessReport` functionality with two new functions designed for enterprise access review workflows:

1. **Create-PerOwnerReports** - Generates individual reports for each share owner
2. **Prepare-OwnerConfirmationEmail** - Creates professional email confirmations in Microsoft Outlook

These functions enable automated, scalable access review processes suitable for corporate and banking audit requirements.

## Requirements

### PowerShell Version
- PowerShell 5.1 or higher

### Required Functions
- `Generate-ShareAccessReport` function (from Generate-ShareAccessReport.ps1)

### Optional Modules
```powershell
# For Excel report generation
Install-Module -Name ImportExcel -Scope CurrentUser -Force

# For PDF report generation
Install-Module -Name PSWritePDF -Scope CurrentUser -Force
```

### For Email Functionality
- Microsoft Outlook installed and configured
- Outlook COM object access (typically requires Outlook to be run at least once)

## Functions

### Create-PerOwnerReports

Generates separate share access reports for each unique owner.

#### Synopsis
Takes expanded share access data and creates individual reports for each owner, filtering to only include shares they own (either as Owner1 or Owner2).

#### Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| Data | PSCustomObject[] | Yes | Array of share access data objects |
| OutputDirectory | String | No | Directory for reports (Default: current directory) |
| Formats | String[] | No | Report formats to generate: "HTML", "XLSX", "PDF" (Default: all) |
| Theme | String | No | Visual theme (Default: "CorporateBlue") |
| Expandable | Switch | No | Create collapsible report sections |

#### Returns
Array of PSCustomObjects with generation results for each owner:
- Owner name
- Record count
- Unique shares count
- Report file paths
- Generation timestamp

#### Example Usage

```powershell
# Load required functions
. .\Generate-ShareAccessReport.ps1
. .\ShareAccessReport-Extensions.ps1

# Generate reports for all owners
$results = Create-PerOwnerReports `
    -Data $ExpandedData `
    -OutputDirectory "C:\Reports\Owners" `
    -Formats @("HTML", "XLSX") `
    -Theme "ExecutiveGreen" `
    -Expandable

# View results
$results | Format-Table Owner, RecordCount, UniqueShares
```

#### Features

**Automatic Owner Detection**
- Identifies unique owners from both Owner1 and Owner2 fields
- Treats co-owners equally (both receive full reports for shared ownership)

**Individual Filtering**
- Each owner receives only their relevant data
- Shares owned by the person in either Owner1 or Owner2 role

**Flexible Output**
- Choose specific formats (HTML, XLSX, PDF)
- Timestamped filenames prevent overwrites
- Sanitized filenames handle special characters

**Error Handling**
- Graceful handling of owners with no shares
- Detailed error messages for troubleshooting
- Continues processing if individual reports fail

### Prepare-OwnerConfirmationEmail

Creates professional email confirmations in Microsoft Outlook for share access review.

#### Synopsis
Prepares HTML-formatted emails for each owner requesting review and confirmation of access rights to their shares. Emails are displayed in Outlook but not sent automatically, allowing manual review.

#### Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| Data | PSCustomObject[] | Yes | Array of share access data |
| OwnerEmails | Hashtable | No | Owner name to email mapping |
| DeadlineDate | DateTime | No | Review deadline (Default: 14 days) |
| Signature | String | No | HTML signature block |
| AttachReports | Switch | No | Attach HTML reports |
| ReportsDirectory | String | No | Location of reports to attach |
| CompanyName | String | No | Company name for branding |
| ContactEmail | String | No | Support contact email |
| SubjectPrefix | String | No | Email subject prefix (Default: "Action Required") |

#### Returns
PSCustomObject with email creation statistics:
- Total owners processed
- Emails created count
- Deadline date
- Creation timestamp

#### Example Usage

```powershell
# Define owner email addresses
$ownerEmails = @{
    "Jane Doe" = "jane.doe@company.com"
    "John Smith" = "john.smith@company.com"
}

# Create emails without attachments
Prepare-OwnerConfirmationEmail `
    -Data $ExpandedData `
    -OwnerEmails $ownerEmails `
    -DeadlineDate (Get-Date).AddDays(14) `
    -CompanyName "Contoso Corporation" `
    -ContactEmail "it-security@contoso.com"
```

#### Example with Report Attachments

```powershell
# First, generate individual reports
$reportResults = Create-PerOwnerReports `
    -Data $ExpandedData `
    -OutputDirectory "C:\Reports\Owners" `
    -Formats @("HTML")

# Then create emails with attachments
Prepare-OwnerConfirmationEmail `
    -Data $ExpandedData `
    -OwnerEmails $ownerEmails `
    -AttachReports `
    -ReportsDirectory "C:\Reports\Owners" `
    -DeadlineDate (Get-Date).AddDays(21) `
    -CompanyName "Contoso Corporation" `
    -ContactEmail "compliance@contoso.com"
```

#### Features

**Professional Email Design**
- Corporate-styled HTML formatting
- Responsive layout suitable for all email clients
- Professional color scheme matching banking/corporate standards
- Clear section headers and formatting

**Comprehensive Content**
- Personalized greeting with owner name
- Complete list of owned shares (server, share name, path)
- Detailed access rights by AD group
- User listings with domain, username, display name, department
- Color-coded risk indicators (high/medium/low)
- Clear action items and deadlines

**Flexible Configuration**
- Customizable company branding
- Configurable deadline dates
- Custom signature blocks
- Support contact information

**Safety Features**
- Emails displayed but NOT sent automatically
- Requires manual review and send
- Missing email addresses handled gracefully
- Reports Outlook connection issues clearly

## Complete Workflow Example

### Step 1: Run ACL Scan
```powershell
# Use your existing ACL scanning script
# This populates $ExpandedData
```

### Step 2: Generate Standard Reports (Optional)
```powershell
. .\Generate-ShareAccessReport.ps1

$timestamp = Get-Date -Format "yyyyMMdd_HHmm"
$reportsPath = "C:\Reports"

# Overall server report
Generate-ShareAccessReport `
    -Data $ExpandedData `
    -ReportType "PerServer" `
    -Theme "CorporateBlue" `
    -HtmlPath "$reportsPath\AllServers_$timestamp.html" `
    -XlsxPath "$reportsPath\AllServers_$timestamp.xlsx"
```

### Step 3: Generate Per-Owner Reports
```powershell
. .\ShareAccessReport-Extensions.ps1

$ownerReports = Create-PerOwnerReports `
    -Data $ExpandedData `
    -OutputDirectory "$reportsPath\PerOwner" `
    -Formats @("HTML", "XLSX") `
    -Theme "ExecutiveGreen" `
    -Expandable
```

### Step 4: Prepare Owner Emails
```powershell
# Define email mappings
$ownerEmails = @{
    "Owner1" = "owner1@company.com"
    "Owner2" = "owner2@company.com"
    # ... add all owners
}

# Custom signature
$signature = @"
<div style="margin-top: 30px; border-top: 1px solid #ccc; padding-top: 20px;">
    <p><strong>Your Name</strong><br>
    IT Security Manager<br>
    Company Name<br>
    Phone: (555) 123-4567<br>
    Email: security@company.com</p>
</div>
"@

# Create emails
$emailResults = Prepare-OwnerConfirmationEmail `
    -Data $ExpandedData `
    -OwnerEmails $ownerEmails `
    -DeadlineDate (Get-Date).AddDays(14) `
    -Signature $signature `
    -CompanyName "Your Company" `
    -ContactEmail "it-support@company.com" `
    -AttachReports `
    -ReportsDirectory "$reportsPath\PerOwner"
```

### Step 5: Review and Send
1. Review each email in Outlook
2. Verify recipient addresses
3. Check attachments
4. Send manually

## Data Structure

Input data must be an array of PSCustomObjects with these properties:

```powershell
[PSCustomObject]@{
    Server      = "ServerName"          # File server name
    Share       = "ShareName"           # Share name
    SharePath   = "D:\Path\To\Share"    # Physical path
    Owner1      = "Primary Owner"       # Primary share owner
    Owner2      = "Secondary Owner"     # Secondary owner (can be $null)
    ADGroupName = "GroupName"           # Active Directory group name
    Domain      = "DomainName"          # User domain
    User        = "username"            # User account name
    DisplayName = "Full Name"           # User display name
    UserGroup   = "DepartmentName"      # User's group/department
    Rights      = "FullControl"         # Access rights
}
```

## Email Content Structure

### Email Sections

1. **Header**
   - Professional branding
   - Company name and department

2. **Greeting**
   - Personalized with owner name
   - Context for the review

3. **Action Required Box**
   - Highlighted section with deadline
   - Clear call to action

4. **Your Owned Shares**
   - Table listing all shares owned
   - Server, share name, and path

5. **Access Rights Summary**
   - Detailed breakdown by share
   - Grouped by AD group
   - User listings with all details
   - Color-coded risk indicators

6. **What You Need To Do**
   - Numbered action items
   - Clear instructions
   - Response options

7. **How to Respond**
   - Contact information
   - Deadline reminder

8. **Signature**
   - Customizable signature block
   - Contact details

## Troubleshooting

### Outlook Not Installed
```
Error: Failed to create Outlook COM object
```
**Solution**: Install Microsoft Outlook or run on a machine with Outlook installed.

### Outlook Permission Issues
```
Error: Access denied to Outlook COM object
```
**Solution**: 
1. Run Outlook at least once manually
2. Ensure PowerShell has appropriate permissions
3. Check Outlook security settings

### Missing Email Addresses
```
Warning: No email address found for 'Owner Name'
```
**Solution**: Either:
- Add owner to OwnerEmails hashtable
- Manually enter email in Outlook when displayed

### Report Attachment Not Found
```
Warning: No report found for owner 'Name'
```
**Solution**:
- Ensure Create-PerOwnerReports was run first
- Verify ReportsDirectory path is correct
- Check that report files exist

### Performance with Many Owners
For organizations with many share owners (>50):
- Process in batches
- Use `-Verbose` to monitor progress
- Consider scheduling during off-hours

## Best Practices

### Email Management
1. **Review Before Sending**: Always review emails before sending
2. **Test First**: Send test emails to yourself first
3. **Batch Processing**: Process owners in manageable groups
4. **Track Responses**: Maintain a spreadsheet of owners and response status

### Report Organization
1. **Dedicated Directory**: Use dedicated output directories
2. **Timestamps**: Keep timestamped reports for audit trail
3. **Archive Old Reports**: Move old reports to archive after review cycle
4. **Access Control**: Ensure report directories have appropriate permissions

### Email Content
1. **Clear Deadlines**: Set realistic deadlines (recommend 14-21 days)
2. **Professional Tone**: Keep language formal and clear
3. **Contact Information**: Provide clear support contacts
4. **Follow-up Plan**: Have a process for non-responders

### Security Considerations
1. **Sensitive Data**: Reports contain sensitive access information
2. **Secure Transmission**: Consider encrypted email if required
3. **Retention Policy**: Define report retention periods
4. **Access Logging**: Log who receives reports and when

## Integration with Existing Scripts

### With Get-ACLScan2
```powershell
# After running Get-ACLScan2.ps1, you have $ExpandedData

# Load extension functions
. .\Generate-ShareAccessReport.ps1
. .\ShareAccessReport-Extensions.ps1

# Generate owner reports
Create-PerOwnerReports -Data $ExpandedData -OutputDirectory "Reports\Owners"

# Create emails
Prepare-OwnerConfirmationEmail -Data $ExpandedData -OwnerEmails $emailMap
```

### Scheduled Execution
```powershell
# Create a scheduled task script
$scriptPath = "C:\Scripts\ShareAccessReview.ps1"
$trigger = New-JobTrigger -Weekly -DaysOfWeek Monday -At "8:00AM"
Register-ScheduledJob -Name "ShareAccessReview" -FilePath $scriptPath -Trigger $trigger
```

## Version History

- **Version 1.0** (Initial Release)
  - Create-PerOwnerReports function
  - Prepare-OwnerConfirmationEmail function
  - Professional HTML email templates
  - Comprehensive error handling
  - Example scripts and documentation

## License

This extension is provided as-is for use in enterprise environments. Modify as needed for your specific requirements.

## Support

For issues or questions:
1. Review this documentation
2. Check the example script for usage patterns
3. Test with sample data first
4. Review function comments and verbose output

## Related Files

- `Generate-ShareAccessReport.ps1` - Core reporting function
- `Generate-ShareAccessReport-README.md` - Core function documentation
- `ShareAccessReport-Extensions-Example.ps1` - Usage examples
- `Get-ACLScan`, `Get-ACLScan2` - ACL scanning scripts that provide input data

## Author

Enterprise PowerShell Team - Specialized in creating professional reporting and automation tools for corporate environments.
