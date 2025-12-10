<#
.SYNOPSIS
    Example usage of Generate-ShareAccessReport function

.DESCRIPTION
    This script demonstrates how to use the Generate-ShareAccessReport function
    with sample data. It shows various usage scenarios and parameter combinations.

.NOTES
    Before running, ensure the following modules are installed:
    - ImportExcel: Install-Module -Name ImportExcel -Scope CurrentUser
    - PSWritePDF: Install-Module -Name PSWritePDF -Scope CurrentUser
#>

# Import the function
. "$PSScriptRoot\Generate-ShareAccessReport.ps1"

#region Sample Data Creation

Write-Host "Creating sample data..." -ForegroundColor Cyan

# Create sample expanded data that matches the expected structure
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
        Owner1 = "Jane Doe"
        Owner2 = "Mike Johnson"
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
        Owner1 = "Jane Doe"
        Owner2 = "Mike Johnson"
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
        Owner1 = "Jane Doe"
        Owner2 = "Mike Johnson"
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
        Owner1 = "Sarah Thompson"
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
        Owner1 = "Sarah Thompson"
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
        Owner1 = "David Lee"
        Owner2 = "Alice Williams"
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
        Owner1 = "David Lee"
        Owner2 = "Alice Williams"
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
        Owner1 = "David Lee"
        Owner2 = "Alice Williams"
        Rights = "FullControl"
    }
    [PSCustomObject]@{
        Server = "FileServer02"
        Share = "Marketing"
        ADGroupName = "MKT_Team"
        Domain = "CORP"
        User = "tmartinez"
        DisplayName = "Thomas Martinez"
        UserGroup = "Marketing"
        SharePath = "E:\Shares\Marketing"
        Owner1 = "Thomas Martinez"
        Owner2 = $null
        Rights = "Modify, Write"
    }
    [PSCustomObject]@{
        Server = "FileServer02"
        Share = "Marketing"
        ADGroupName = "MKT_ReadOnly"
        Domain = "CORP"
        User = "landerson"
        DisplayName = "Linda Anderson"
        UserGroup = "Marketing Support"
        SharePath = "E:\Shares\Marketing"
        Owner1 = "Thomas Martinez"
        Owner2 = $null
        Rights = "Read"
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
        Owner1 = "Jennifer Thomas"
        Owner2 = "Kevin Moore"
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
        Owner1 = "Jennifer Thomas"
        Owner2 = "Kevin Moore"
        Rights = "FullControl"
    }
    [PSCustomObject]@{
        Server = "FileServer03"
        Share = "IT"
        ADGroupName = "IT_Admins"
        Domain = "CORP"
        User = "cjackson"
        DisplayName = "Christopher Jackson"
        UserGroup = "IT Administration"
        SharePath = "F:\Shares\IT"
        Owner1 = "Christopher Jackson"
        Owner2 = $null
        Rights = "FullControl"
    }
    [PSCustomObject]@{
        Server = "FileServer03"
        Share = "IT"
        ADGroupName = "IT_Support"
        Domain = "CORP"
        User = "mwhite"
        DisplayName = "Michelle White"
        UserGroup = "IT Support"
        SharePath = "F:\Shares\IT"
        Owner1 = "Christopher Jackson"
        Owner2 = $null
        Rights = "Modify, Write"
    }
    [PSCustomObject]@{
        Server = "FileServer03"
        Share = "IT"
        ADGroupName = "IT_ReadOnly"
        Domain = "CORP"
        User = "dharris"
        DisplayName = "Daniel Harris"
        UserGroup = "IT Support"
        SharePath = "F:\Shares\IT"
        Owner1 = "Christopher Jackson"
        Owner2 = $null
        Rights = "Read"
    }
)

Write-Host "Sample data created: $($SampleExpandedData.Count) records" -ForegroundColor Green
Write-Host ""

#endregion

#region Example 1: Basic PerOwner Report with CorporateBlue Theme

Write-Host "========================================" -ForegroundColor Magenta
Write-Host "Example 1: PerOwner Report - CorporateBlue Theme" -ForegroundColor Magenta
Write-Host "========================================" -ForegroundColor Magenta
Write-Host ""

try {
    $result1 = Generate-ShareAccessReport `
        -Data $SampleExpandedData `
        -ReportType "PerOwner" `
        -Theme "CorporateBlue" `
        -HtmlPath "$PSScriptRoot\Example1_PerOwner_CorporateBlue.html" `
        -XlsxPath "$PSScriptRoot\Example1_PerOwner_CorporateBlue.xlsx" `
        -PdfPath "$PSScriptRoot\Example1_PerOwner_CorporateBlue.pdf" `
        -Verbose

    Write-Host "`nExample 1 Summary:" -ForegroundColor Yellow
    $result1.Summary | Format-List
}
catch {
    Write-Error "Example 1 failed: $($_.Exception.Message)"
}

Write-Host ""

#endregion

#region Example 2: Expandable PerServer Report with MinimalGray Theme

Write-Host "========================================" -ForegroundColor Magenta
Write-Host "Example 2: PerServer Report - MinimalGray Theme (Expandable)" -ForegroundColor Magenta
Write-Host "========================================" -ForegroundColor Magenta
Write-Host ""

try {
    $result2 = Generate-ShareAccessReport `
        -Data $SampleExpandedData `
        -ReportType "PerServer" `
        -Theme "MinimalGray" `
        -Expandable `
        -HtmlPath "$PSScriptRoot\Example2_PerServer_MinimalGray_Expandable.html" `
        -XlsxPath "$PSScriptRoot\Example2_PerServer_MinimalGray_Expandable.xlsx" `
        -PdfPath "$PSScriptRoot\Example2_PerServer_MinimalGray_Expandable.pdf" `
        -Verbose

    Write-Host "`nExample 2 Summary:" -ForegroundColor Yellow
    $result2.Summary | Format-List
}
catch {
    Write-Error "Example 2 failed: $($_.Exception.Message)"
}

Write-Host ""

#endregion

#region Example 3: PerOwner Report with ExecutiveGreen Theme

Write-Host "========================================" -ForegroundColor Magenta
Write-Host "Example 3: PerOwner Report - ExecutiveGreen Theme" -ForegroundColor Magenta
Write-Host "========================================" -ForegroundColor Magenta
Write-Host ""

try {
    $result3 = Generate-ShareAccessReport `
        -Data $SampleExpandedData `
        -ReportType "PerOwner" `
        -Theme "ExecutiveGreen" `
        -HtmlPath "$PSScriptRoot\Example3_PerOwner_ExecutiveGreen.html" `
        -XlsxPath "$PSScriptRoot\Example3_PerOwner_ExecutiveGreen.xlsx" `
        -PdfPath "$PSScriptRoot\Example3_PerOwner_ExecutiveGreen.pdf" `
        -Verbose

    Write-Host "`nExample 3 Summary:" -ForegroundColor Yellow
    $result3.Summary | Format-List
}
catch {
    Write-Error "Example 3 failed: $($_.Exception.Message)"
}

Write-Host ""

#endregion

#region Example 4: Test with Empty Data

Write-Host "========================================" -ForegroundColor Magenta
Write-Host "Example 4: Empty Data Handling" -ForegroundColor Magenta
Write-Host "========================================" -ForegroundColor Magenta
Write-Host ""

try {
    $result4 = Generate-ShareAccessReport `
        -Data @() `
        -ReportType "PerServer" `
        -Theme "CorporateBlue" `
        -HtmlPath "$PSScriptRoot\Example4_EmptyData.html" `
        -Verbose

    Write-Host "`nExample 4 Summary:" -ForegroundColor Yellow
    $result4.Summary | Format-List
}
catch {
    Write-Error "Example 4 failed: $($_.Exception.Message)"
}

Write-Host ""

#endregion

#region Example 5: Using Default Paths (Timestamped)

Write-Host "========================================" -ForegroundColor Magenta
Write-Host "Example 5: Using Default Timestamped Paths" -ForegroundColor Magenta
Write-Host "========================================" -ForegroundColor Magenta
Write-Host ""

try {
    $result5 = Generate-ShareAccessReport `
        -Data $SampleExpandedData `
        -ReportType "PerServer" `
        -Theme "ExecutiveGreen" `
        -Expandable `
        -Verbose

    Write-Host "`nExample 5 Generated Files:" -ForegroundColor Yellow
    Write-Host "HTML: $($result5.HtmlReport)" -ForegroundColor Cyan
    Write-Host "XLSX: $($result5.XlsxReport)" -ForegroundColor Cyan
    Write-Host "PDF:  $($result5.PdfReport)" -ForegroundColor Cyan
}
catch {
    Write-Error "Example 5 failed: $($_.Exception.Message)"
}

Write-Host ""

#endregion

#region Integration Example with Existing Code Pattern

Write-Host "========================================" -ForegroundColor Magenta
Write-Host "Integration Example: Mimicking Repository Pattern" -ForegroundColor Magenta
Write-Host "========================================" -ForegroundColor Magenta
Write-Host ""

<#
This demonstrates how to integrate the function with the existing code pattern
seen in the repository (similar to Get-ACLScan2 script):

# After collecting ACL data and expanding groups (from existing scripts):
# $ExpandedData = @() ... (populated from your scan)

# Generate reports in all formats
$reportResult = Generate-ShareAccessReport `
    -Data $ExpandedData `
    -ReportType "PerServer" `
    -Theme "CorporateBlue" `
    -Expandable `
    -HtmlPath "$ReportsPath\ShareAccess_PerServer_$Timestamp.html" `
    -XlsxPath "$ReportsPath\ShareAccess_PerServer_$Timestamp.xlsx" `
    -PdfPath "$ReportsPath\ShareAccess_PerServer_$Timestamp.pdf"

# Also generate per-owner view
$reportResult = Generate-ShareAccessReport `
    -Data $ExpandedData `
    -ReportType "PerOwner" `
    -Theme "ExecutiveGreen" `
    -Expandable `
    -HtmlPath "$ReportsPath\ShareAccess_PerOwner_$Timestamp.html" `
    -XlsxPath "$ReportsPath\ShareAccess_PerOwner_$Timestamp.xlsx" `
    -PdfPath "$ReportsPath\ShareAccess_PerOwner_$Timestamp.pdf"

# Open the reports folder
Invoke-Item $ReportsPath
#>

Write-Host "See comments in script for integration example" -ForegroundColor Green
Write-Host ""

#endregion

Write-Host "========================================" -ForegroundColor Green
Write-Host "All examples completed!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "Generated files are in: $PSScriptRoot" -ForegroundColor Cyan
Write-Host ""
Write-Host "To install required modules if not already installed:" -ForegroundColor Yellow
Write-Host "  Install-Module -Name ImportExcel -Scope CurrentUser -Force" -ForegroundColor Gray
Write-Host "  Install-Module -Name PSWritePDF -Scope CurrentUser -Force" -ForegroundColor Gray
