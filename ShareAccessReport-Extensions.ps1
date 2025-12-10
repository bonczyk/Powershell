<#
.SYNOPSIS
    Extended functions for Generate-ShareAccessReport to create per-owner reports and email confirmations.

.DESCRIPTION
    This module extends the Generate-ShareAccessReport functionality with:
    - Create-PerOwnerReports: Generates separate reports for each owner
    - Prepare-OwnerConfirmationEmail: Creates Outlook email confirmations for each owner

.NOTES
    Author: Enterprise PowerShell Team
    Version: 1.0
    Requires: Generate-ShareAccessReport.ps1, Microsoft Outlook (for email functionality)
#>

#region Helper Functions for AD User Lookup

<#
.SYNOPSIS
    Resolves email addresses to display names using Active Directory.

.DESCRIPTION
    Takes an email address and queries Active Directory to retrieve the user's display name.
    Caches results to improve performance for repeated lookups.

.PARAMETER EmailAddress
    Email address to resolve.

.PARAMETER Cache
    Optional hashtable for caching results to improve performance.

.RETURNS
    Display name from AD, or the original email if not found or AD unavailable.
#>
function Get-DisplayNameFromEmail {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$EmailAddress,

        [Parameter(Mandatory = $false)]
        [hashtable]$Cache = @{}
    )

    # Return null/empty if input is null/empty
    if ([string]::IsNullOrWhiteSpace($EmailAddress)) {
        return $null
    }

    # Check cache first
    if ($Cache.ContainsKey($EmailAddress)) {
        return $Cache[$EmailAddress]
    }

    try {
        # Try to query Active Directory
        $adUser = Get-ADUser -Filter "EmailAddress -eq '$EmailAddress'" -Properties DisplayName -ErrorAction Stop
        
        if ($adUser -and $adUser.DisplayName) {
            $displayName = $adUser.DisplayName
            Write-Verbose "Resolved '$EmailAddress' to '$displayName'"
        }
        else {
            # AD user found but no display name, use the Name property
            if ($adUser -and $adUser.Name) {
                $displayName = $adUser.Name
                Write-Verbose "Resolved '$EmailAddress' to '$displayName' (using Name)"
            }
            else {
                # Not found, use email as fallback
                $displayName = $EmailAddress
                Write-Verbose "Could not resolve '$EmailAddress', using email as display name"
            }
        }
    }
    catch {
        # AD query failed (module not available, no connection, etc.)
        $displayName = $EmailAddress
        Write-Verbose "Failed to query AD for '$EmailAddress': $($_.Exception.Message). Using email as display name."
    }

    # Cache the result
    $Cache[$EmailAddress] = $displayName
    
    return $displayName
}

<#
.SYNOPSIS
    Enriches data with display names from email addresses.

.DESCRIPTION
    Takes the expanded data and resolves Owner1 and Owner2 email addresses to display names,
    adding OwnerDisplayName1 and OwnerDisplayName2 properties.

.PARAMETER Data
    Array of PSCustomObjects with Owner1 and Owner2 email addresses.

.RETURNS
    Enriched data with display name properties added.
#>
function Add-OwnerDisplayNames {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject[]]$Data
    )

    Write-Verbose "Enriching data with owner display names from Active Directory..."
    
    # Create a cache for performance
    $displayNameCache = @{}
    
    # Get unique email addresses to resolve
    $uniqueEmails = @($Data | ForEach-Object { $_.Owner1; $_.Owner2 } | Where-Object { $_ } | Select-Object -Unique)
    
    Write-Host "Resolving $($uniqueEmails.Count) unique owner email(s) to display names..." -ForegroundColor Cyan
    
    # Pre-populate cache
    foreach ($email in $uniqueEmails) {
        $null = Get-DisplayNameFromEmail -EmailAddress $email -Cache $displayNameCache
    }
    
    # Add display name properties to each record
    $enrichedData = foreach ($record in $Data) {
        # Create a copy with new properties
        $enriched = $record.PSObject.Copy()
        
        # Add display names
        if ($record.Owner1) {
            Add-Member -InputObject $enriched -NotePropertyName 'OwnerDisplayName1' -NotePropertyValue $displayNameCache[$record.Owner1] -Force
        }
        if ($record.Owner2) {
            Add-Member -InputObject $enriched -NotePropertyName 'OwnerDisplayName2' -NotePropertyValue $displayNameCache[$record.Owner2] -Force
        }
        
        $enriched
    }
    
    Write-Host "✓ Owner display names resolved" -ForegroundColor Green
    Write-Host ""
    
    return $enrichedData
}

#endregion

#region Create-PerOwnerReports Function

<#
.SYNOPSIS
    Generates separate share access reports for each unique owner.

.DESCRIPTION
    Takes the expanded share access data and creates individual reports for each owner,
    filtering to only include shares they own (either as Owner1 or Owner2). Each owner
    gets their own set of reports in the specified formats.
    
    Owner1 and Owner2 are treated as email addresses and will be resolved to display names
    using Active Directory (Get-ADUser). Display names are shown in reports along with email addresses.

.PARAMETER Data
    Mandatory. Array of PSCustomObjects containing share access information with properties:
    Server, Share, ADGroupName, Domain, User, DisplayName, UserGroup, SharePath, Owner1, Owner2, Rights
    Note: Owner1 and Owner2 should be email addresses. Display names will be resolved from Active Directory.

.PARAMETER OutputDirectory
    Directory where owner reports will be saved. Defaults to current directory.

.PARAMETER Formats
    Array of report formats to generate. Valid values: "HTML", "XLSX", "PDF"
    Defaults to all three formats.

.PARAMETER Theme
    HTML theme for styling. Valid values: "CorporateBlue", "MinimalGray", "ExecutiveGreen"
    Defaults to "CorporateBlue".

.PARAMETER Expandable
    Switch parameter. When specified, creates collapsible/expandable report sections.

.EXAMPLE
    Create-PerOwnerReports -Data $ExpandedData -OutputDirectory "C:\Reports\Owners"

.EXAMPLE
    Create-PerOwnerReports -Data $ExpandedData -Formats @("HTML", "XLSX") -Theme "ExecutiveGreen" -Expandable

.NOTES
    Requires the Generate-ShareAccessReport function to be loaded.
#>
function Create-PerOwnerReports {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $false)]
        [ValidateNotNullOrEmpty()]
        [PSCustomObject[]]$Data,

        [Parameter(Mandatory = $false)]
        [string]$OutputDirectory = $PWD,

        [Parameter(Mandatory = $false)]
        [ValidateSet("HTML", "XLSX", "PDF")]
        [string[]]$Formats = @("HTML", "XLSX", "PDF"),

        [Parameter(Mandatory = $false)]
        [ValidateSet("CorporateBlue", "MinimalGray", "ExecutiveGreen")]
        [string]$Theme = "CorporateBlue",

        [Parameter(Mandatory = $false)]
        [switch]$Expandable
    )

    begin {
        # Ensure output directory exists
        if (-not (Test-Path $OutputDirectory)) {
            Write-Verbose "Creating output directory: $OutputDirectory"
            New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null
        }

        # Validate that Generate-ShareAccessReport is available
        if (-not (Get-Command Generate-ShareAccessReport -ErrorAction SilentlyContinue)) {
            throw "Generate-ShareAccessReport function is not available. Please load Generate-ShareAccessReport.ps1 first."
        }

        Write-Verbose "Output Directory: $OutputDirectory"
        Write-Verbose "Formats: $($Formats -join ', ')"
        Write-Verbose "Theme: $Theme"
    }

    process {
        # Enrich data with display names from email addresses
        $enrichedData = Add-OwnerDisplayNames -Data $Data
        
        # Get unique owner emails (combining Owner1 and Owner2)
        Write-Verbose "Identifying unique owners..."
        $uniqueOwnerEmails = @($enrichedData | ForEach-Object { $_.Owner1; $_.Owner2 } | Where-Object { $_ } | Select-Object -Unique | Sort-Object)

        if ($uniqueOwnerEmails.Count -eq 0) {
            Write-Warning "No owners found in the provided data."
            return
        }

        # Create mapping of emails to display names
        $emailToDisplayName = @{}
        foreach ($record in $enrichedData) {
            if ($record.Owner1 -and -not $emailToDisplayName.ContainsKey($record.Owner1)) {
                $emailToDisplayName[$record.Owner1] = $record.OwnerDisplayName1
            }
            if ($record.Owner2 -and -not $emailToDisplayName.ContainsKey($record.Owner2)) {
                $emailToDisplayName[$record.Owner2] = $record.OwnerDisplayName2
            }
        }

        Write-Host "Found $($uniqueOwnerEmails.Count) unique owner(s)" -ForegroundColor Cyan
        Write-Host ""

        $results = @()
        $ownerIndex = 0

        foreach ($ownerEmail in $uniqueOwnerEmails) {
            $ownerIndex++
            $ownerDisplayName = $emailToDisplayName[$ownerEmail]
            Write-Host "[$ownerIndex/$($uniqueOwnerEmails.Count)] Processing owner: $ownerDisplayName ($ownerEmail)" -ForegroundColor Yellow

            # Filter data for this owner (shares they own as Owner1 or Owner2)
            $ownerData = $enrichedData | Where-Object { $_.Owner1 -eq $ownerEmail -or $_.Owner2 -eq $ownerEmail }

            if ($ownerData.Count -eq 0) {
                Write-Warning "  No shares found for owner: $ownerDisplayName"
                continue
            }

            Write-Verbose "  Found $($ownerData.Count) record(s) for $ownerDisplayName"

            # Sanitize display name for file naming
            $safeOwnerName = $ownerDisplayName -replace '[\\/:*?"<>|]', '_'
            $safeOwnerName = $safeOwnerName -replace '\s+', '_'

            # Build paths for each format
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $htmlPath = $null
            $xlsxPath = $null
            $pdfPath = $null

            if ("HTML" -in $Formats) {
                $htmlPath = Join-Path $OutputDirectory "ShareAccessReport_${safeOwnerName}_${timestamp}.html"
            }
            if ("XLSX" -in $Formats) {
                $xlsxPath = Join-Path $OutputDirectory "ShareAccessReport_${safeOwnerName}_${timestamp}.xlsx"
            }
            if ("PDF" -in $Formats) {
                $pdfPath = Join-Path $OutputDirectory "ShareAccessReport_${safeOwnerName}_${timestamp}.pdf"
            }

            # Generate the report for this owner
            try {
                $reportParams = @{
                    Data       = $ownerData
                    ReportType = "PerOwner"
                    Theme      = $Theme
                }

                if ($htmlPath) { $reportParams['HtmlPath'] = $htmlPath }
                if ($xlsxPath) { $reportParams['XlsxPath'] = $xlsxPath }
                if ($pdfPath) { $reportParams['PdfPath'] = $pdfPath }
                if ($Expandable) { $reportParams['Expandable'] = $true }

                $result = Generate-ShareAccessReport @reportParams

                # Store result information
                $results += [PSCustomObject]@{
                    OwnerDisplayName = $ownerDisplayName
                    OwnerEmail       = $ownerEmail
                    RecordCount      = $ownerData.Count
                    UniqueShares     = ($ownerData.Share | Select-Object -Unique).Count
                    HtmlReport       = $htmlPath
                    XlsxReport       = $xlsxPath
                    PdfReport        = $pdfPath
                    GeneratedDate    = Get-Date
                }

                Write-Host "  ✓ Report generated successfully" -ForegroundColor Green
                if ($htmlPath) { Write-Host "    HTML: $htmlPath" -ForegroundColor Gray }
                if ($xlsxPath) { Write-Host "    XLSX: $xlsxPath" -ForegroundColor Gray }
                if ($pdfPath) { Write-Host "    PDF:  $pdfPath" -ForegroundColor Gray }
            }
            catch {
                Write-Error "  Failed to generate report for $ownerDisplayName : $($_.Exception.Message)"
            }

            Write-Host ""
        }

        Write-Host "========================================" -ForegroundColor Green
        Write-Host "Per-Owner Report Generation Complete" -ForegroundColor Green
        Write-Host "========================================" -ForegroundColor Green
        Write-Host "Total Owners Processed: $($results.Count)" -ForegroundColor Cyan
        Write-Host ""

        return $results
    }
}

#endregion

#region Prepare-OwnerConfirmationEmail Function

<#
.SYNOPSIS
    Prepares Outlook email confirmations for each share owner to review access rights.

.DESCRIPTION
    Creates professional, HTML-formatted emails in Microsoft Outlook for each owner,
    asking them to review and confirm access rights to their shares. Each email includes
    a summary of shares, AD groups with access, and detailed user listings. The function
    can attach individual owner reports and displays emails for manual review before sending.

.PARAMETER Data
    Mandatory. Array of PSCustomObjects containing share access information with properties:
    Server, Share, ADGroupName, Domain, User, DisplayName, UserGroup, SharePath, Owner1, Owner2, Rights

.PARAMETER OwnerEmails
    Optional hashtable mapping owner names to email addresses.
    Example: @{ "John Doe" = "john.doe@company.com"; "Jane Smith" = "jane.smith@company.com" }
    If not provided, will try to use owner name as email or prompt for manual entry in Outlook.

.PARAMETER DeadlineDate
    Deadline date for access review confirmation. Defaults to 14 days from now.

.PARAMETER Signature
    Email signature block to append to each email. Should be HTML formatted.
    Defaults to a generic corporate signature.

.PARAMETER AttachReports
    Switch parameter. When specified, attaches the HTML report for each owner.
    Requires that reports have been generated first (e.g., via Create-PerOwnerReports).

.PARAMETER ReportsDirectory
    Directory where owner reports are located (for attachments). Defaults to current directory.

.PARAMETER CompanyName
    Company name to use in email branding. Defaults to "Our Organization".

.PARAMETER ContactEmail
    Contact email for questions. Defaults to "it-support@company.com".

.PARAMETER SubjectPrefix
    Optional prefix for email subject. Defaults to "Action Required".

.EXAMPLE
    Prepare-OwnerConfirmationEmail -Data $ExpandedData -OwnerEmails $emailMap

.EXAMPLE
    $emails = @{ "John Doe" = "john.doe@company.com" }
    Prepare-OwnerConfirmationEmail -Data $ExpandedData -OwnerEmails $emails `
        -DeadlineDate (Get-Date).AddDays(7) -AttachReports -ReportsDirectory "C:\Reports\Owners"

.NOTES
    Requires Microsoft Outlook to be installed and configured.
    Emails are displayed but not sent automatically - requires manual review and send.
#>
function Prepare-OwnerConfirmationEmail {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $false)]
        [ValidateNotNullOrEmpty()]
        [PSCustomObject[]]$Data,

        [Parameter(Mandatory = $false)]
        [hashtable]$OwnerEmails = @{},

        [Parameter(Mandatory = $false)]
        [datetime]$DeadlineDate = (Get-Date).AddDays(14),

        [Parameter(Mandatory = $false)]
        [string]$Signature = "",

        [Parameter(Mandatory = $false)]
        [switch]$AttachReports,

        [Parameter(Mandatory = $false)]
        [string]$ReportsDirectory = $PWD,

        [Parameter(Mandatory = $false)]
        [string]$CompanyName = "Our Organization",

        [Parameter(Mandatory = $false)]
        [string]$ContactEmail = "it-support@company.com",

        [Parameter(Mandatory = $false)]
        [string]$SubjectPrefix = "Action Required"
    )

    begin {
        # Try to create Outlook application object
        try {
            Write-Verbose "Attempting to connect to Microsoft Outlook..."
            $outlook = New-Object -ComObject Outlook.Application
            Write-Host "✓ Successfully connected to Microsoft Outlook" -ForegroundColor Green
        }
        catch {
            throw "Failed to create Outlook COM object. Please ensure Microsoft Outlook is installed and configured. Error: $($_.Exception.Message)"
        }

        # Set default signature if not provided
        if ([string]::IsNullOrWhiteSpace($Signature)) {
            $Signature = @"
<div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #ccc; font-size: 0.9em; color: #666;">
    <p><strong>IT Security & Compliance Team</strong><br>
    $CompanyName<br>
    For questions or concerns, please contact: <a href="mailto:$ContactEmail">$ContactEmail</a></p>
</div>
"@
        }

        Write-Verbose "Company Name: $CompanyName"
        Write-Verbose "Contact Email: $ContactEmail"
        Write-Verbose "Deadline Date: $($DeadlineDate.ToString('MMMM dd, yyyy'))"
    }

    process {
        # Enrich data with display names from email addresses
        $enrichedData = Add-OwnerDisplayNames -Data $Data
        
        # Get unique owner emails (Owner1 and Owner2 are email addresses)
        Write-Verbose "Identifying unique owners..."
        $uniqueOwnerEmails = @($enrichedData | ForEach-Object { $_.Owner1; $_.Owner2 } | Where-Object { $_ } | Select-Object -Unique | Sort-Object)

        if ($uniqueOwnerEmails.Count -eq 0) {
            Write-Warning "No owners found in the provided data."
            return
        }

        # Create mapping of emails to display names
        $emailToDisplayName = @{}
        foreach ($record in $enrichedData) {
            if ($record.Owner1 -and -not $emailToDisplayName.ContainsKey($record.Owner1)) {
                $emailToDisplayName[$record.Owner1] = $record.OwnerDisplayName1
            }
            if ($record.Owner2 -and -not $emailToDisplayName.ContainsKey($record.Owner2)) {
                $emailToDisplayName[$record.Owner2] = $record.OwnerDisplayName2
            }
        }

        Write-Host ""
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host "Preparing Confirmation Emails" -ForegroundColor Cyan
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host "Found $($uniqueOwnerEmails.Count) unique owner(s)" -ForegroundColor Cyan
        Write-Host ""

        $emailsCreated = 0
        $ownerIndex = 0

        foreach ($ownerEmail in $uniqueOwnerEmails) {
            $ownerIndex++
            $ownerDisplayName = $emailToDisplayName[$ownerEmail]
            Write-Host "[$ownerIndex/$($uniqueOwnerEmails.Count)] Preparing email for: $ownerDisplayName ($ownerEmail)" -ForegroundColor Yellow

            # Owner1/Owner2 are already email addresses, use them directly
            # OwnerEmails parameter is deprecated but kept for backward compatibility
            $recipientEmail = $ownerEmail
            if ($OwnerEmails.ContainsKey($ownerDisplayName)) {
                # Legacy: if display name mapping exists, use it
                $recipientEmail = $OwnerEmails[$ownerDisplayName]
                Write-Verbose "  Using email from OwnerEmails parameter: $recipientEmail"
            }

            # Filter data for this owner
            $ownerData = $enrichedData | Where-Object { $_.Owner1 -eq $ownerEmail -or $_.Owner2 -eq $ownerEmail }

            if ($ownerData.Count -eq 0) {
                Write-Warning "  No shares found for owner: $ownerDisplayName"
                continue
            }

            # Get unique shares for this owner
            $ownerShares = $ownerData | Select-Object Server, Share, SharePath -Unique | Sort-Object Server, Share

            # Create email body using display name
            $emailBody = Build-ConfirmationEmailBody -Owner $ownerDisplayName -OwnerEmail $ownerEmail `
                -OwnerData $ownerData -OwnerShares $ownerShares `
                -DeadlineDate $DeadlineDate -CompanyName $CompanyName -ContactEmail $ContactEmail -Signature $Signature

            # Create new email
            try {
                $mail = $outlook.CreateItem(0) # 0 = olMailItem

                # Set email properties - use the email address as recipient
                if ($recipientEmail) {
                    $mail.To = $recipientEmail
                }
                $mail.Subject = "${SubjectPrefix}: Access Review for Your Owned File Shares"
                $mail.HTMLBody = $emailBody

                # Attach report if requested
                if ($AttachReports) {
                    $safeOwnerName = $ownerDisplayName -replace '[\\/:*?"<>|]', '_'
                    $safeOwnerName = $safeOwnerName -replace '\s+', '_'
                    
                    # Find the most recent report for this owner
                    $reportFiles = Get-ChildItem -Path $ReportsDirectory -Filter "ShareAccessReport_${safeOwnerName}_*.html" -ErrorAction SilentlyContinue |
                        Sort-Object LastWriteTime -Descending

                    if ($reportFiles -and $reportFiles.Count -gt 0) {
                        $reportPath = $reportFiles[0].FullName
                        $mail.Attachments.Add($reportPath) | Out-Null
                        Write-Host "  ✓ Attached report: $($reportFiles[0].Name)" -ForegroundColor Gray
                    }
                    else {
                        Write-Warning "  No report found for owner '$ownerDisplayName' in directory: $ReportsDirectory"
                    }
                }

                # Display the email (does not send automatically)
                $mail.Display()

                $emailsCreated++
                Write-Host "  ✓ Email displayed in Outlook (ready for review)" -ForegroundColor Green
            }
            catch {
                Write-Error "  Failed to create email for $ownerDisplayName : $($_.Exception.Message)"
            }

            Write-Host ""
        }

        Write-Host "========================================" -ForegroundColor Green
        Write-Host "Email Preparation Complete" -ForegroundColor Green
        Write-Host "========================================" -ForegroundColor Green
        Write-Host "Emails Created: $emailsCreated" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "NOTE: All emails are displayed in Outlook but NOT sent automatically." -ForegroundColor Yellow
        Write-Host "Please review each email carefully before sending." -ForegroundColor Yellow
        Write-Host ""

        return [PSCustomObject]@{
            TotalOwners     = $uniqueOwners.Count
            EmailsCreated   = $emailsCreated
            DeadlineDate    = $DeadlineDate
            CreatedDate     = Get-Date
        }
    }
}

#endregion

#region Helper Functions

<#
.SYNOPSIS
    Builds the HTML body for owner confirmation email.

.DESCRIPTION
    Internal helper function to create professional, corporate-styled HTML email body
    with share information, access details, and review instructions.
#>
function Build-ConfirmationEmailBody {
    [CmdletBinding()]
    param(
        [string]$Owner,
        [string]$OwnerEmail,
        [PSCustomObject[]]$OwnerData,
        [PSCustomObject[]]$OwnerShares,
        [datetime]$DeadlineDate,
        [string]$CompanyName,
        [string]$ContactEmail,
        [string]$Signature
    )

    $deadlineStr = $DeadlineDate.ToString('MMMM dd, yyyy')
    $shareCount = $OwnerShares.Count
    $totalAccessCount = $OwnerData.Count

    # Group data by share for detailed listing
    $shareGroups = $OwnerData | Group-Object -Property Server, Share

    $emailBody = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <style>
        body {
            font-family: 'Segoe UI', 'Calibri', 'Arial', sans-serif;
            line-height: 1.6;
            color: #333;
            max-width: 800px;
            margin: 0 auto;
            padding: 20px;
            background-color: #f5f5f5;
        }
        .email-container {
            background-color: white;
            padding: 30px;
            border-radius: 5px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.1);
        }
        .header {
            border-bottom: 3px solid #0066cc;
            padding-bottom: 15px;
            margin-bottom: 25px;
        }
        .header h1 {
            color: #0066cc;
            margin: 0;
            font-size: 24px;
        }
        .header .subtitle {
            color: #666;
            font-size: 14px;
            margin-top: 5px;
        }
        .section {
            margin-bottom: 25px;
        }
        .section h2 {
            color: #0066cc;
            font-size: 18px;
            margin-bottom: 10px;
            border-bottom: 1px solid #ddd;
            padding-bottom: 5px;
        }
        .highlight-box {
            background-color: #fff9e6;
            border-left: 4px solid #ffcc00;
            padding: 15px;
            margin: 15px 0;
        }
        .deadline {
            color: #d9534f;
            font-weight: bold;
            font-size: 16px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 10px 0;
            font-size: 14px;
        }
        th {
            background-color: #0066cc;
            color: white;
            padding: 10px;
            text-align: left;
            font-weight: 600;
        }
        td {
            padding: 8px 10px;
            border-bottom: 1px solid #ddd;
        }
        tr:nth-child(even) {
            background-color: #f9f9f9;
        }
        .share-header {
            background-color: #e6f2ff;
            padding: 10px;
            margin: 15px 0 5px 0;
            border-left: 4px solid #0066cc;
            font-weight: bold;
        }
        .high-risk {
            color: #d9534f;
            font-weight: bold;
        }
        .medium-risk {
            color: #f0ad4e;
            font-weight: bold;
        }
        ul {
            margin: 10px 0;
            padding-left: 25px;
        }
        li {
            margin: 5px 0;
        }
        .action-items {
            background-color: #e6f7ff;
            border: 1px solid #91d5ff;
            padding: 15px;
            border-radius: 5px;
            margin: 20px 0;
        }
        .action-items h3 {
            margin-top: 0;
            color: #0066cc;
        }
    </style>
</head>
<body>
    <div class="email-container">
        <div class="header">
            <h1>File Share Access Review Required</h1>
            <div class="subtitle">$CompanyName - IT Security & Compliance</div>
        </div>

        <div class="section">
            <p>Dear <strong>$Owner</strong>,</p>
            
            <p>As part of our ongoing security and compliance initiatives, we are conducting a comprehensive review 
            of access rights to network file shares. Our records indicate that you are listed as an owner for 
            <strong>$shareCount</strong> file share(s) within our organization.</p>
        </div>

        <div class="highlight-box">
            <p><strong>Action Required:</strong> Please review the access permissions listed below and confirm 
            that all users and groups have appropriate access to these shares.</p>
            <p class="deadline">Response Deadline: $deadlineStr</p>
        </div>

        <div class="section">
            <h2>Your Owned Shares</h2>
            <p>You are responsible for the following file shares:</p>
            <table>
                <thead>
                    <tr>
                        <th>Server</th>
                        <th>Share Name</th>
                        <th>Path</th>
                    </tr>
                </thead>
                <tbody>
"@

    # Add share list
    foreach ($share in $OwnerShares) {
        $emailBody += @"
                    <tr>
                        <td>$($share.Server)</td>
                        <td>$($share.Share)</td>
                        <td>$($share.SharePath)</td>
                    </tr>
"@
    }

    $emailBody += @"
                </tbody>
            </table>
        </div>

        <div class="section">
            <h2>Access Rights Summary</h2>
            <p>The following Active Directory groups and users have access to your shares:</p>
"@

    # Add detailed access information grouped by share
    foreach ($shareGroup in $shareGroups) {
        $server = $shareGroup.Values[0]
        $share = $shareGroup.Values[1]
        $shareData = $shareGroup.Group
        
        $emailBody += @"
            <div class="share-header">📁 \\$server\$share</div>
"@

        # Group by AD Group within this share
        $adGroups = $shareData | Group-Object -Property ADGroupName

        foreach ($adGroup in $adGroups) {
            $groupName = $adGroup.Name
            $groupData = $adGroup.Group
            $rights = $groupData[0].Rights

            # Determine risk class
            $riskClass = ""
            if ($rights -match "FullControl") {
                $riskClass = "high-risk"
            }
            elseif ($rights -match "Modify|Write") {
                $riskClass = "medium-risk"
            }

            $emailBody += @"
            <table style="margin-top: 10px;">
                <thead>
                    <tr>
                        <th colspan="4">👥 AD Group: $groupName <span style="float:right;" class="$riskClass">Rights: $rights</span></th>
                    </tr>
                    <tr>
                        <th>Domain</th>
                        <th>Username</th>
                        <th>Display Name</th>
                        <th>Department/Group</th>
                    </tr>
                </thead>
                <tbody>
"@

            foreach ($user in $groupData) {
                $emailBody += @"
                    <tr>
                        <td>$($user.Domain)</td>
                        <td>$($user.User)</td>
                        <td>$($user.DisplayName)</td>
                        <td>$($user.UserGroup)</td>
                    </tr>
"@
            }

            $emailBody += @"
                </tbody>
            </table>
"@
        }
    }

    $emailBody += @"
        </div>

        <div class="action-items">
            <h3>What You Need To Do:</h3>
            <ol>
                <li><strong>Review</strong> the access permissions listed above carefully</li>
                <li><strong>Verify</strong> that all users and groups should have access to your shares</li>
                <li><strong>Identify</strong> any users or groups that should NOT have access</li>
                <li><strong>Respond</strong> by <span class="deadline">$deadlineStr</span> with one of the following:
                    <ul>
                        <li><strong>Confirmation:</strong> All access rights are appropriate and approved</li>
                        <li><strong>Changes Needed:</strong> List any users/groups that should be removed or added</li>
                    </ul>
                </li>
            </ol>
        </div>

        <div class="section">
            <h2>How to Respond</h2>
            <p>Please reply to this email with your confirmation or requested changes. If you have any questions 
            or concerns, please contact us at <a href="mailto:$ContactEmail">$ContactEmail</a>.</p>
            
            <p><strong>Note:</strong> If we do not receive a response by the deadline, we will assume the current 
            access permissions are approved and compliant.</p>
        </div>

        <div class="section">
            <p>Thank you for your cooperation in maintaining the security and compliance of our IT systems.</p>
            <p>Best regards,</p>
        </div>

        $Signature
    </div>
</body>
</html>
"@

    return $emailBody
}

#endregion

# Note: To use as a module, create a .psm1 file and add:
# Export-ModuleMember -Function Create-PerOwnerReports, Prepare-OwnerConfirmationEmail
