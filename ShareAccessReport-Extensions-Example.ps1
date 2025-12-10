<#
.SYNOPSIS
    Example usage of ShareAccessReport-Extensions functions

.DESCRIPTION
    This script demonstrates how to use the Create-PerOwnerReports and 
    Prepare-OwnerConfirmationEmail functions to generate per-owner reports
    and create email confirmations for access review.

.NOTES
    Before running, ensure:
    - ImportExcel module is installed (optional): Install-Module -Name ImportExcel -Scope CurrentUser
    - PSWritePDF module is installed (optional): Install-Module -Name PSWritePDF -Scope CurrentUser
    - Microsoft Outlook is installed and configured (for email functionality)
#>

# Import the required functions
. "$PSScriptRoot\Generate-ShareAccessReport.ps1"
. "$PSScriptRoot\ShareAccessReport-Extensions.ps1"

#region Sample Data Creation

Write-Host "Creating sample data..." -ForegroundColor Cyan

# Create sample expanded data that matches the expected structure
# Note: Owner1 and Owner2 are now email addresses that will be resolved to display names via AD
$SampleExpandedData = @(
    [PSCustomObject]@{
        Server = "FileServer01"
        Share = "Finance"
        ADGroupName = "FIN_ReadOnly"
        Domain = "CORP"
        User = "jsmith"
        DisplayName = "John Smith"
        UserGroup = "Finance Department"
        SharePath = "D:\Shares\Finance"
        Owner1 = "jane.doe@company.com"
        Owner2 = "mike.johnson@company.com"
        Rights = "Read, Execute"
    }
    [PSCustomObject]@{
        Server = "FileServer01"
        Share = "Finance"
        ADGroupName = "FIN_FullAccess"
        Domain = "CORP"
        User = "mjohnson"
        DisplayName = "Mike Johnson"
        UserGroup = "Finance Management"
        SharePath = "D:\Shares\Finance"
        Owner1 = "jane.doe@company.com"
        Owner2 = "mike.johnson@company.com"
        Rights = "FullControl"
    }
    [PSCustomObject]@{
        Server = "FileServer01"
        Share = "Finance"
        ADGroupName = "FIN_FullAccess"
        Domain = "CORP"
        User = "jdoe"
        DisplayName = "Jane Doe"
        UserGroup = "Finance Management"
        SharePath = "D:\Shares\Finance"
        Owner1 = "jane.doe@company.com"
        Owner2 = "mike.johnson@company.com"
        Rights = "FullControl"
    }
    [PSCustomObject]@{
        Server = "FileServer01"
        Share = "HR"
        ADGroupName = "HR_Modify"
        Domain = "CORP"
        User = "sthompson"
        DisplayName = "Sarah Thompson"
        UserGroup = "Human Resources"
        SharePath = "D:\Shares\HR"
        Owner1 = "sarah.thompson@company.com"
        Owner2 = $null
        Rights = "Modify, Write"
    }
    [PSCustomObject]@{
        Server = "FileServer01"
        Share = "HR"
        ADGroupName = "HR_ReadOnly"
        Domain = "CORP"
        User = "rbrown"
        DisplayName = "Robert Brown"
        UserGroup = "HR Support"
        SharePath = "D:\Shares\HR"
        Owner1 = "sarah.thompson@company.com"
        Owner2 = $null
        Rights = "Read"
    }
    [PSCustomObject]@{
        Server = "FileServer02"
        Share = "Engineering"
        ADGroupName = "ENG_Developers"
        Domain = "CORP"
        User = "awilliams"
        DisplayName = "Alice Williams"
        UserGroup = "Engineering"
        SharePath = "E:\Shares\Engineering"
        Owner1 = "david.lee@company.com"
        Owner2 = "alice.williams@company.com"
        Rights = "Modify, Write"
    }
    [PSCustomObject]@{
        Server = "FileServer02"
        Share = "Engineering"
        ADGroupName = "ENG_Developers"
        Domain = "CORP"
        User = "dlee"
        DisplayName = "David Lee"
        UserGroup = "Engineering"
        SharePath = "E:\Shares\Engineering"
        Owner1 = "david.lee@company.com"
        Owner2 = "alice.williams@company.com"
        Rights = "Modify, Write"
    }
    [PSCustomObject]@{
        Server = "FileServer02"
        Share = "Engineering"
        ADGroupName = "ENG_Architects"
        Domain = "CORP"
        User = "pgarcia"
        DisplayName = "Patricia Garcia"
        UserGroup = "Engineering Architecture"
        SharePath = "E:\Shares\Engineering"
        Owner1 = "david.lee@company.com"
        Owner2 = "alice.williams@company.com"
        Rights = "FullControl"
    }
    [PSCustomObject]@{
        Server = "FileServer03"
        Share = "Legal"
        ADGroupName = "LEG_FullAccess"
        Domain = "CORP"
        User = "jthomas"
        DisplayName = "Jennifer Thomas"
        UserGroup = "Legal Department"
        SharePath = "F:\Shares\Legal"
        Owner1 = "jennifer.thomas@company.com"
        Owner2 = "kevin.moore@company.com"
        Rights = "FullControl"
    }
    [PSCustomObject]@{
        Server = "FileServer03"
        Share = "Legal"
        ADGroupName = "LEG_FullAccess"
        Domain = "CORP"
        User = "kmoore"
        DisplayName = "Kevin Moore"
        UserGroup = "Legal Department"
        SharePath = "F:\Shares\Legal"
        Owner1 = "jennifer.thomas@company.com"
        Owner2 = "kevin.moore@company.com"
        Rights = "FullControl"
    }
)

Write-Host "NOTE: Owner1 and Owner2 are email addresses. In a real AD environment," -ForegroundColor Yellow
Write-Host "      these will be resolved to display names automatically." -ForegroundColor Yellow

Write-Host "Sample data created: $($SampleExpandedData.Count) records" -ForegroundColor Green
Write-Host ""

#endregion

#region Example 1: Create Per-Owner Reports

Write-Host "========================================" -ForegroundColor Magenta
Write-Host "Example 1: Create Per-Owner Reports" -ForegroundColor Magenta
Write-Host "========================================" -ForegroundColor Magenta
Write-Host ""

try {
    # Create output directory for owner reports
    $ownerReportsDir = Join-Path $PSScriptRoot "OwnerReports"
    
    Write-Host "Generating individual reports for each owner..." -ForegroundColor Cyan
    
    $reportResults = Create-PerOwnerReports `
        -Data $SampleExpandedData `
        -OutputDirectory $ownerReportsDir `
        -Formats @("HTML", "XLSX") `
        -Theme "CorporateBlue" `
        -Expandable `
        -Verbose

    Write-Host "Report Generation Results:" -ForegroundColor Yellow
    $reportResults | Format-Table OwnerDisplayName, OwnerEmail, RecordCount, UniqueShares -AutoSize
    
    Write-Host "All owner reports have been generated in: $ownerReportsDir" -ForegroundColor Green
}
catch {
    Write-Error "Example 1 failed: $($_.Exception.Message)"
}

Write-Host ""

#endregion

#region Example 2: Prepare Owner Confirmation Emails (Without Attachments)

Write-Host "========================================" -ForegroundColor Magenta
Write-Host "Example 2: Prepare Confirmation Emails" -ForegroundColor Magenta
Write-Host "========================================" -ForegroundColor Magenta
Write-Host ""

try {
    # Note: OwnerEmails parameter is now optional since Owner1/Owner2 are email addresses
    # It's kept for backward compatibility if you need to override email addresses
    # In this case, we don't need it since Owner1/Owner2 are already emails
    $ownerEmails = @{}  # Empty - will use Owner1/Owner2 as email addresses directly

    Write-Host "NOTE: This example will attempt to open Microsoft Outlook." -ForegroundColor Yellow
    Write-Host "If Outlook is not installed, this will fail gracefully." -ForegroundColor Yellow
    Write-Host ""
    
    $response = Read-Host "Do you want to create Outlook emails? (Y/N)"
    
    if ($response -eq 'Y' -or $response -eq 'y') {
        Write-Host "Creating confirmation emails in Outlook..." -ForegroundColor Cyan
        
        $emailResults = Prepare-OwnerConfirmationEmail `
            -Data $SampleExpandedData `
            -OwnerEmails $ownerEmails `
            -DeadlineDate (Get-Date).AddDays(14) `
            -CompanyName "Contoso Corporation" `
            -ContactEmail "it-security@contoso.com" `
            -SubjectPrefix "Action Required" `
            -Verbose

        Write-Host "Email Preparation Results:" -ForegroundColor Yellow
        $emailResults | Format-List
    }
    else {
        Write-Host "Skipping email creation." -ForegroundColor Gray
    }
}
catch {
    Write-Warning "Example 2 failed (this is expected if Outlook is not installed): $($_.Exception.Message)"
}

Write-Host ""

#endregion

#region Example 3: Complete Workflow - Reports + Emails with Attachments

Write-Host "========================================" -ForegroundColor Magenta
Write-Host "Example 3: Complete Workflow with Attachments" -ForegroundColor Magenta
Write-Host "========================================" -ForegroundColor Magenta
Write-Host ""

try {
    # Step 1: Generate reports
    Write-Host "Step 1: Generating per-owner reports..." -ForegroundColor Cyan
    
    $workflowReportsDir = Join-Path $PSScriptRoot "WorkflowReports"
    
    $reportResults = Create-PerOwnerReports `
        -Data $SampleExpandedData `
        -OutputDirectory $workflowReportsDir `
        -Formats @("HTML") `
        -Theme "ExecutiveGreen" `
        -Verbose

    Write-Host "✓ Reports generated" -ForegroundColor Green
    Write-Host ""

    # Step 2: Create emails with attachments
    Write-Host "Step 2: Preparing emails with report attachments..." -ForegroundColor Cyan
    
    $ownerEmails = @{
        "Jane Doe" = "jane.doe@company.com"
        "Mike Johnson" = "mike.johnson@company.com"
        "Sarah Thompson" = "sarah.thompson@company.com"
        "David Lee" = "david.lee@company.com"
        "Alice Williams" = "alice.williams@company.com"
        "Jennifer Thomas" = "jennifer.thomas@company.com"
        "Kevin Moore" = "kevin.moore@company.com"
    }
    
    $response = Read-Host "Do you want to create Outlook emails with attachments? (Y/N)"
    
    if ($response -eq 'Y' -or $response -eq 'y') {
        $emailResults = Prepare-OwnerConfirmationEmail `
            -Data $SampleExpandedData `
            -OwnerEmails $ownerEmails `
            -DeadlineDate (Get-Date).AddDays(21) `
            -CompanyName "Contoso Corporation" `
            -ContactEmail "compliance@contoso.com" `
            -AttachReports `
            -ReportsDirectory $workflowReportsDir `
            -Verbose

        Write-Host "✓ Emails created with attachments" -ForegroundColor Green
    }
    else {
        Write-Host "Skipping email creation." -ForegroundColor Gray
    }
}
catch {
    Write-Warning "Example 3 failed: $($_.Exception.Message)"
}

Write-Host ""

#endregion

#region Integration Example with Existing ACL Scanning

Write-Host "========================================" -ForegroundColor Magenta
Write-Host "Integration Example: With ACL Scanning" -ForegroundColor Magenta
Write-Host "========================================" -ForegroundColor Magenta
Write-Host ""

<#
This demonstrates the complete workflow integrating with existing ACL scanning:

# Step 1: Run your existing ACL scan (Get-ACLScan or Get-ACLScan2)
# This populates $ExpandedData with the share access information

# Step 2: Generate standard reports (optional)
. .\Generate-ShareAccessReport.ps1
$timestamp = Get-Date -Format "yyyyMMdd_HHmm"
$reportsPath = "C:\Reports"

# Generate overall reports
Generate-ShareAccessReport `
    -Data $ExpandedData `
    -ReportType "PerServer" `
    -Theme "CorporateBlue" `
    -HtmlPath "$reportsPath\AllServers_$timestamp.html" `
    -XlsxPath "$reportsPath\AllServers_$timestamp.xlsx"

Generate-ShareAccessReport `
    -Data $ExpandedData `
    -ReportType "PerOwner" `
    -Theme "ExecutiveGreen" `
    -HtmlPath "$reportsPath\AllOwners_$timestamp.html" `
    -XlsxPath "$reportsPath\AllOwners_$timestamp.xlsx"

# Step 3: Generate individual owner reports
. .\ShareAccessReport-Extensions.ps1

$ownerReports = Create-PerOwnerReports `
    -Data $ExpandedData `
    -OutputDirectory "$reportsPath\PerOwner" `
    -Formats @("HTML", "XLSX", "PDF") `
    -Theme "CorporateBlue" `
    -Expandable

# Step 4: Prepare confirmation emails for owners
$ownerEmails = @{
    # Map owner names to email addresses
    "John Doe" = "john.doe@company.com"
    "Jane Smith" = "jane.smith@company.com"
    # ... add all owners
}

$emailResults = Prepare-OwnerConfirmationEmail `
    -Data $ExpandedData `
    -OwnerEmails $ownerEmails `
    -DeadlineDate (Get-Date).AddDays(14) `
    -CompanyName "Your Company Name" `
    -ContactEmail "it-security@company.com" `
    -AttachReports `
    -ReportsDirectory "$reportsPath\PerOwner"

# Step 5: Review emails in Outlook and send manually
Write-Host "Review and send the emails from Outlook"

# Step 6: Open reports folder
Invoke-Item $reportsPath
#>

Write-Host "See comments in script for complete integration workflow" -ForegroundColor Green
Write-Host ""

#endregion

Write-Host "========================================" -ForegroundColor Green
Write-Host "All examples completed!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "Note: Owner reports and emails are created in subdirectories:" -ForegroundColor Cyan
Write-Host "  - OwnerReports\" -ForegroundColor Gray
Write-Host "  - WorkflowReports\" -ForegroundColor Gray
Write-Host ""
