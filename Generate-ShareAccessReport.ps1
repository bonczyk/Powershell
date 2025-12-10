<#
.SYNOPSIS
    Generates professional share access reports in multiple formats (HTML, XLSX, PDF).

.DESCRIPTION
    This function takes expanded share access data and generates professional, corporate-quality 
    reports in HTML, Excel (XLSX), and PDF formats. Reports can be organized by Owner or Server,
    with optional expandable/collapsible hierarchies and multiple visual themes.

.PARAMETER Data
    Mandatory. Array of PSCustomObjects containing share access information with properties:
    Server, Share, ADGroupName, Domain, User, DisplayName, UserGroup, SharePath, Owner1, Owner2, Rights

.PARAMETER ReportType
    Type of report to generate. Valid values: "PerOwner", "PerServer"
    - PerOwner: Groups data by Owner1/Owner2
    - PerServer: Groups data by Server

.PARAMETER Expandable
    Switch parameter. When specified, creates collapsible/expandable report sections using HTML details/summary tags.

.PARAMETER Theme
    HTML theme for styling. Valid values: "CorporateBlue", "MinimalGray", "ExecutiveGreen"

.PARAMETER HtmlPath
    Output path for HTML report. If not specified, uses current directory with timestamped filename.

.PARAMETER XlsxPath
    Output path for Excel report. If not specified, uses current directory with timestamped filename.

.PARAMETER PdfPath
    Output path for PDF report. If not specified, uses current directory with timestamped filename.

.EXAMPLE
    Generate-ShareAccessReport -Data $ExpandedData -ReportType PerOwner -Theme CorporateBlue

.EXAMPLE
    Generate-ShareAccessReport -Data $ExpandedData -ReportType PerServer -Expandable -Theme MinimalGray `
        -HtmlPath "C:\Reports\ShareAccess.html" -XlsxPath "C:\Reports\ShareAccess.xlsx"

.NOTES
    Author: Enterprise PowerShell Team
    Version: 1.0
    Requires: ImportExcel module, PSWritePDF module
#>

function Generate-ShareAccessReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $false)]
        [AllowEmptyCollection()]
        [PSCustomObject[]]$Data,

        [Parameter(Mandatory = $true)]
        [ValidateSet("PerOwner", "PerServer")]
        [string]$ReportType,

        [Parameter(Mandatory = $false)]
        [switch]$Expandable,

        [Parameter(Mandatory = $false)]
        [ValidateSet("CorporateBlue", "MinimalGray", "ExecutiveGreen")]
        [string]$Theme = "CorporateBlue",

        [Parameter(Mandatory = $false)]
        [string]$HtmlPath,

        [Parameter(Mandatory = $false)]
        [string]$XlsxPath,

        [Parameter(Mandatory = $false)]
        [string]$PdfPath
    )

    begin {
        # Validate that required modules are available
        if (-not (Get-Module -ListAvailable -Name 'ImportExcel')) {
            Write-Warning "Module 'ImportExcel' is not installed. Excel (XLSX) report generation will be skipped."
        }
        if (-not (Get-Module -ListAvailable -Name 'PSWritePDF')) {
            Write-Warning "Module 'PSWritePDF' is not installed. PDF report generation will be skipped."
        }

        # Set default paths if not specified
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        if (-not $HtmlPath) {
            $HtmlPath = Join-Path $PWD "ShareAccessReport_${ReportType}_${timestamp}.html"
        }
        if (-not $XlsxPath) {
            $XlsxPath = Join-Path $PWD "ShareAccessReport_${ReportType}_${timestamp}.xlsx"
        }
        if (-not $PdfPath) {
            $PdfPath = Join-Path $PWD "ShareAccessReport_${ReportType}_${timestamp}.pdf"
        }

        Write-Verbose "Report Type: $ReportType"
        Write-Verbose "HTML Output: $HtmlPath"
        Write-Verbose "XLSX Output: $XlsxPath"
        Write-Verbose "PDF Output: $PdfPath"
    }

    process {
        # Handle empty data
        if ($null -eq $Data -or $Data.Count -eq 0) {
            Write-Warning "No data provided. Generating empty report with warning message."
            $Data = @()
        }

        # Calculate summary statistics
        $summary = Get-ReportSummary -Data $Data

        # Generate HTML Report
        Write-Verbose "Generating HTML report..."
        $htmlContent = Generate-HtmlReport -Data $Data -ReportType $ReportType -Theme $Theme -Expandable:$Expandable -Summary $summary
        $htmlContent | Out-File -FilePath $HtmlPath -Encoding UTF8 -Force
        Write-Host "HTML report generated: $HtmlPath" -ForegroundColor Green

        # Generate XLSX Report
        if (Get-Module -ListAvailable -Name ImportExcel) {
            Write-Verbose "Generating Excel report..."
            Generate-XlsxReport -Data $Data -ReportType $ReportType -OutputPath $XlsxPath -Summary $summary
            Write-Host "Excel report generated: $XlsxPath" -ForegroundColor Green
        }

        # Generate PDF Report
        if (Get-Module -ListAvailable -Name PSWritePDF) {
            Write-Verbose "Generating PDF report..."
            Generate-PdfReport -HtmlPath $HtmlPath -OutputPath $PdfPath -ReportType $ReportType
            Write-Host "PDF report generated: $PdfPath" -ForegroundColor Green
        }

        # Return summary information
        return [PSCustomObject]@{
            HtmlReport = $HtmlPath
            XlsxReport = $XlsxPath
            PdfReport  = $PdfPath
            Summary    = $summary
        }
    }
}

#region Helper Functions

function Get-ReportSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [PSCustomObject[]]$Data
    )

    if ($Data.Count -eq 0) {
        return [PSCustomObject]@{
            TotalRecords    = 0
            UniqueServers   = 0
            UniqueShares    = 0
            UniqueOwners    = 0
            UniqueUsers     = 0
            UniqueADGroups  = 0
            HighRiskCount   = 0
        }
    }

    # Efficiently collect unique owners
    $uniqueOwners = @($Data | ForEach-Object { $_.Owner1; $_.Owner2 } | Where-Object { $_ } | Select-Object -Unique)

    return [PSCustomObject]@{
        TotalRecords    = $Data.Count
        UniqueServers   = ($Data.Server | Select-Object -Unique).Count
        UniqueShares    = ($Data.Share | Select-Object -Unique).Count
        UniqueOwners    = $uniqueOwners.Count
        UniqueUsers     = ($Data.User | Where-Object { $_ } | Select-Object -Unique).Count
        UniqueADGroups  = ($Data.ADGroupName | Select-Object -Unique).Count
        HighRiskCount   = ($Data | Where-Object { $_.Rights -match "FullControl|Modify" }).Count
    }
}

function Get-ThemeStyles {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("CorporateBlue", "MinimalGray", "ExecutiveGreen")]
        [string]$Theme
    )

    switch ($Theme) {
        "CorporateBlue" {
            return @"
        body { 
            font-family: 'Segoe UI', 'Arial', sans-serif; 
            background-color: #f0f4f8; 
            color: #2c3e50; 
            margin: 0; 
            padding: 20px;
            line-height: 1.6;
        }
        .header { 
            background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
            color: white; 
            padding: 30px; 
            text-align: center; 
            border-radius: 8px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            margin-bottom: 30px;
        }
        .header h1 { 
            margin: 0 0 10px 0; 
            font-size: 2.5em;
            font-weight: 300;
            letter-spacing: 1px;
        }
        .header .subtitle { 
            font-size: 1.1em; 
            opacity: 0.9;
            font-weight: 300;
        }
        .summary-section {
            background: white;
            border-radius: 8px;
            padding: 25px;
            margin-bottom: 30px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.08);
            border-left: 5px solid #2a5298;
        }
        .summary-section h2 {
            color: #1e3c72;
            margin-top: 0;
            font-size: 1.8em;
            border-bottom: 2px solid #e0e6ed;
            padding-bottom: 10px;
        }
        .summary-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-top: 20px;
        }
        .summary-item {
            background: #f8fafc;
            padding: 15px;
            border-radius: 6px;
            border: 1px solid #e0e6ed;
        }
        .summary-label {
            font-size: 0.85em;
            color: #64748b;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 5px;
        }
        .summary-value {
            font-size: 2em;
            font-weight: 600;
            color: #1e3c72;
        }
        .group-container { 
            background: white; 
            border: 1px solid #cbd5e1; 
            margin-bottom: 20px; 
            border-radius: 8px; 
            overflow: hidden;
            box-shadow: 0 2px 4px rgba(0,0,0,0.06);
        }
        .group-header { 
            background: linear-gradient(to right, #2a5298, #1e3c72);
            color: white; 
            padding: 15px 20px; 
            cursor: pointer; 
            font-weight: 600;
            font-size: 1.1em;
            display: flex;
            align-items: center;
            transition: background 0.3s ease;
        }
        .group-header:hover {
            background: linear-gradient(to right, #1e3c72, #16325c);
        }
        .group-content { 
            padding: 20px;
            background: #fafbfc;
        }
        .subgroup-header {
            background: #e8eef5;
            color: #1e3c72;
            padding: 12px 18px;
            cursor: pointer;
            font-weight: 600;
            border-left: 4px solid #2a5298;
            margin-bottom: 10px;
            transition: background 0.2s ease;
        }
        .subgroup-header:hover {
            background: #dae3f0;
        }
        table { 
            width: 100%; 
            border-collapse: collapse; 
            font-size: 0.95em;
            background: white;
            border-radius: 6px;
            overflow: hidden;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        }
        th { 
            background: #1e3c72;
            color: white; 
            text-align: left; 
            padding: 12px 15px;
            font-weight: 600;
            text-transform: uppercase;
            font-size: 0.85em;
            letter-spacing: 0.5px;
        }
        td { 
            padding: 10px 15px; 
            border-bottom: 1px solid #e0e6ed;
        }
        tr:hover { 
            background-color: #f8fafc;
        }
        tr:last-child td {
            border-bottom: none;
        }
        .rights-high { 
            background-color: #fef2f2; 
            color: #dc2626;
            font-weight: 600;
        }
        .rights-medium {
            background-color: #fffbeb;
            color: #f59e0b;
            font-weight: 600;
        }
        .rights-low {
            background-color: #f0fdf4;
            color: #16a34a;
        }
        .footer {
            text-align: center;
            margin-top: 40px;
            padding: 20px;
            background: white;
            border-radius: 8px;
            color: #64748b;
            font-size: 0.9em;
            box-shadow: 0 2px 4px rgba(0,0,0,0.06);
        }
        details > summary {
            list-style: none;
        }
        details > summary::-webkit-details-marker {
            display: none;
        }
        details > summary::before {
            content: '▶ ';
            margin-right: 8px;
            transition: transform 0.2s ease;
            display: inline-block;
        }
        details[open] > summary::before {
            transform: rotate(90deg);
        }
"@
        }
        "MinimalGray" {
            return @"
        body { 
            font-family: 'Helvetica Neue', 'Arial', sans-serif; 
            background-color: #fafafa; 
            color: #212121; 
            margin: 0; 
            padding: 20px;
            line-height: 1.6;
        }
        .header { 
            background: #212121;
            color: white; 
            padding: 30px; 
            text-align: center; 
            margin-bottom: 30px;
        }
        .header h1 { 
            margin: 0 0 10px 0; 
            font-size: 2.5em;
            font-weight: 300;
        }
        .header .subtitle { 
            font-size: 1.1em; 
            opacity: 0.85;
            font-weight: 300;
        }
        .summary-section {
            background: white;
            padding: 25px;
            margin-bottom: 30px;
            border: 1px solid #e0e0e0;
        }
        .summary-section h2 {
            color: #212121;
            margin-top: 0;
            font-size: 1.8em;
            border-bottom: 1px solid #e0e0e0;
            padding-bottom: 10px;
            font-weight: 400;
        }
        .summary-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-top: 20px;
        }
        .summary-item {
            background: #f5f5f5;
            padding: 15px;
            border: 1px solid #e0e0e0;
        }
        .summary-label {
            font-size: 0.85em;
            color: #757575;
            text-transform: uppercase;
            letter-spacing: 1px;
            margin-bottom: 5px;
        }
        .summary-value {
            font-size: 2em;
            font-weight: 300;
            color: #212121;
        }
        .group-container { 
            background: white; 
            border: 1px solid #e0e0e0; 
            margin-bottom: 20px; 
        }
        .group-header { 
            background: #424242;
            color: white; 
            padding: 15px 20px; 
            cursor: pointer; 
            font-weight: 500;
            font-size: 1.1em;
        }
        .group-header:hover {
            background: #616161;
        }
        .group-content { 
            padding: 20px;
        }
        .subgroup-header {
            background: #f5f5f5;
            color: #212121;
            padding: 12px 18px;
            cursor: pointer;
            font-weight: 500;
            border-left: 3px solid #757575;
            margin-bottom: 10px;
        }
        .subgroup-header:hover {
            background: #e0e0e0;
        }
        table { 
            width: 100%; 
            border-collapse: collapse; 
            font-size: 0.95em;
            background: white;
        }
        th { 
            background: #424242;
            color: white; 
            text-align: left; 
            padding: 12px 15px;
            font-weight: 500;
            text-transform: uppercase;
            font-size: 0.85em;
            letter-spacing: 1px;
        }
        td { 
            padding: 10px 15px; 
            border-bottom: 1px solid #e0e0e0;
        }
        tr:hover { 
            background-color: #fafafa;
        }
        tr:last-child td {
            border-bottom: none;
        }
        .rights-high { 
            background-color: #ffebee; 
            color: #c62828;
            font-weight: 500;
        }
        .rights-medium {
            background-color: #fff3e0;
            color: #f57c00;
            font-weight: 500;
        }
        .rights-low {
            background-color: #e8f5e9;
            color: #2e7d32;
        }
        .footer {
            text-align: center;
            margin-top: 40px;
            padding: 20px;
            background: white;
            color: #757575;
            font-size: 0.9em;
            border-top: 1px solid #e0e0e0;
        }
        details > summary {
            list-style: none;
        }
        details > summary::-webkit-details-marker {
            display: none;
        }
        details > summary::before {
            content: '▸ ';
            margin-right: 8px;
            display: inline-block;
        }
        details[open] > summary::before {
            content: '▾ ';
        }
"@
        }
        "ExecutiveGreen" {
            return @"
        body { 
            font-family: 'Georgia', 'Times New Roman', serif; 
            background-color: #f8faf9; 
            color: #1a3a2a; 
            margin: 0; 
            padding: 20px;
            line-height: 1.7;
        }
        .header { 
            background: linear-gradient(135deg, #2d5a3d 0%, #1a4028 100%);
            color: white; 
            padding: 35px; 
            text-align: center; 
            border-radius: 8px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            margin-bottom: 30px;
        }
        .header h1 { 
            margin: 0 0 10px 0; 
            font-size: 2.8em;
            font-weight: 400;
            letter-spacing: 2px;
        }
        .header .subtitle { 
            font-size: 1.2em; 
            opacity: 0.9;
            font-weight: 300;
            font-style: italic;
        }
        .summary-section {
            background: white;
            border-radius: 8px;
            padding: 30px;
            margin-bottom: 30px;
            box-shadow: 0 2px 8px rgba(45,90,61,0.1);
            border-top: 4px solid #2d5a3d;
        }
        .summary-section h2 {
            color: #1a4028;
            margin-top: 0;
            font-size: 2em;
            border-bottom: 2px solid #d4e4da;
            padding-bottom: 12px;
            font-weight: 500;
        }
        .summary-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
            gap: 25px;
            margin-top: 25px;
        }
        .summary-item {
            background: linear-gradient(to bottom right, #f8faf9, #e8f4ed);
            padding: 20px;
            border-radius: 6px;
            border: 1px solid #c8dbd1;
            box-shadow: 0 2px 4px rgba(45,90,61,0.05);
        }
        .summary-label {
            font-size: 0.9em;
            color: #5a7a65;
            text-transform: uppercase;
            letter-spacing: 1px;
            margin-bottom: 8px;
            font-family: 'Segoe UI', sans-serif;
        }
        .summary-value {
            font-size: 2.2em;
            font-weight: 600;
            color: #1a4028;
        }
        .group-container { 
            background: white; 
            border: 1px solid #c8dbd1; 
            margin-bottom: 25px; 
            border-radius: 8px; 
            overflow: hidden;
            box-shadow: 0 2px 6px rgba(45,90,61,0.08);
        }
        .group-header { 
            background: linear-gradient(to right, #2d5a3d, #1a4028);
            color: white; 
            padding: 18px 22px; 
            cursor: pointer; 
            font-weight: 600;
            font-size: 1.15em;
            display: flex;
            align-items: center;
            transition: background 0.3s ease;
        }
        .group-header:hover {
            background: linear-gradient(to right, #1a4028, #0f2618);
        }
        .group-content { 
            padding: 22px;
            background: #fafcfb;
        }
        .subgroup-header {
            background: #e8f4ed;
            color: #1a4028;
            padding: 14px 20px;
            cursor: pointer;
            font-weight: 600;
            border-left: 5px solid #2d5a3d;
            margin-bottom: 12px;
            border-radius: 0 4px 4px 0;
            transition: all 0.2s ease;
        }
        .subgroup-header:hover {
            background: #d4e4da;
            border-left-width: 8px;
        }
        table { 
            width: 100%; 
            border-collapse: collapse; 
            font-size: 0.95em;
            background: white;
            border-radius: 6px;
            overflow: hidden;
            box-shadow: 0 1px 4px rgba(45,90,61,0.08);
        }
        th { 
            background: #1a4028;
            color: white; 
            text-align: left; 
            padding: 14px 16px;
            font-weight: 600;
            text-transform: uppercase;
            font-size: 0.88em;
            letter-spacing: 0.8px;
            font-family: 'Segoe UI', sans-serif;
        }
        td { 
            padding: 12px 16px; 
            border-bottom: 1px solid #e8f4ed;
        }
        tr:hover { 
            background-color: #f8faf9;
        }
        tr:last-child td {
            border-bottom: none;
        }
        .rights-high { 
            background-color: #fef2f2; 
            color: #991b1b;
            font-weight: 600;
        }
        .rights-medium {
            background-color: #fffbeb;
            color: #b45309;
            font-weight: 600;
        }
        .rights-low {
            background-color: #f0fdf4;
            color: #15803d;
        }
        .footer {
            text-align: center;
            margin-top: 45px;
            padding: 25px;
            background: white;
            border-radius: 8px;
            color: #5a7a65;
            font-size: 0.95em;
            box-shadow: 0 2px 6px rgba(45,90,61,0.08);
            border-top: 3px solid #2d5a3d;
        }
        details > summary {
            list-style: none;
        }
        details > summary::-webkit-details-marker {
            display: none;
        }
        details > summary::before {
            content: '▶ ';
            margin-right: 10px;
            transition: transform 0.2s ease;
            display: inline-block;
        }
        details[open] > summary::before {
            transform: rotate(90deg);
        }
"@
        }
    }
}

function Get-RightsClass {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Rights
    )

    if ($Rights -match "FullControl") {
        return "rights-high"
    }
    elseif ($Rights -match "Modify|Write") {
        return "rights-medium"
    }
    else {
        return "rights-low"
    }
}

function Generate-HtmlReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [PSCustomObject[]]$Data,

        [Parameter(Mandatory = $true)]
        [string]$ReportType,

        [Parameter(Mandatory = $true)]
        [string]$Theme,

        [Parameter(Mandatory = $false)]
        [switch]$Expandable,

        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Summary
    )

    $styles = Get-ThemeStyles -Theme $Theme
    $reportDate = Get-Date -Format "MMMM dd, yyyy HH:mm:ss"
    
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Share Access Report - $ReportType</title>
    <style>
$styles
    </style>
</head>
<body>
    <div class="header">
        <h1>Network Share Access Report</h1>
        <div class="subtitle">$ReportType Analysis | Generated: $reportDate</div>
    </div>

    <div class="summary-section">
        <h2>Executive Summary</h2>
        <div class="summary-grid">
            <div class="summary-item">
                <div class="summary-label">Total Records</div>
                <div class="summary-value">$($Summary.TotalRecords)</div>
            </div>
            <div class="summary-item">
                <div class="summary-label">Unique Servers</div>
                <div class="summary-value">$($Summary.UniqueServers)</div>
            </div>
            <div class="summary-item">
                <div class="summary-label">Unique Shares</div>
                <div class="summary-value">$($Summary.UniqueShares)</div>
            </div>
            <div class="summary-item">
                <div class="summary-label">Unique Owners</div>
                <div class="summary-value">$($Summary.UniqueOwners)</div>
            </div>
            <div class="summary-item">
                <div class="summary-label">Unique Users</div>
                <div class="summary-value">$($Summary.UniqueUsers)</div>
            </div>
            <div class="summary-item">
                <div class="summary-label">AD Groups</div>
                <div class="summary-value">$($Summary.UniqueADGroups)</div>
            </div>
        </div>
    </div>
"@

    if ($Data.Count -eq 0) {
        $html += @"
    <div class="summary-section">
        <h2>No Data Available</h2>
        <p>No share access data was provided for this report.</p>
    </div>
"@
    }
    else {
        if ($ReportType -eq "PerOwner") {
            $html += Generate-PerOwnerHtml -Data $Data -Expandable:$Expandable
        }
        else {
            $html += Generate-PerServerHtml -Data $Data -Expandable:$Expandable
        }
    }

    $html += @"
    <div class="footer">
        <p><strong>Confidential Report - For Internal Use Only</strong></p>
        <p>Generated by Enterprise Share Access Reporting System | Contact: IT Security Department</p>
        <p>&copy; $(Get-Date -Format yyyy) Enterprise IT Security | All Rights Reserved</p>
    </div>
</body>
</html>
"@

    return $html
}

function Generate-PerOwnerHtml {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [PSCustomObject[]]$Data,

        [Parameter(Mandatory = $false)]
        [switch]$Expandable
    )

    $html = ""
    
    # Group by Owner (considering both Owner1 and Owner2)
    $ownerGroups = @{}
    foreach ($item in $Data) {
        $owners = @()
        if ($item.Owner1) { $owners += $item.Owner1 }
        if ($item.Owner2) { $owners += $item.Owner2 }
        
        if ($owners.Count -eq 0) {
            $owners = @("Unassigned")
        }

        foreach ($owner in $owners) {
            if (-not $ownerGroups.ContainsKey($owner)) {
                $ownerGroups[$owner] = @()
            }
            $ownerGroups[$owner] += $item
        }
    }

    $sortedOwners = $ownerGroups.Keys | Sort-Object

    foreach ($owner in $sortedOwners) {
        $ownerData = $ownerGroups[$owner]
        $ownerRecordCount = $ownerData.Count
        
        $detailsOpen = if ($Expandable) { "" } else { "open" }
        
        $html += @"
    <details class="group-container" $detailsOpen>
        <summary class="group-header">👤 Owner: $owner ($ownerRecordCount records)</summary>
        <div class="group-content">
"@

        # Group by Share within Owner
        $shareGroups = $ownerData | Group-Object -Property Share
        
        foreach ($shareGroup in $shareGroups) {
            $shareName = $shareGroup.Name
            $shareData = $shareGroup.Group
            $serverName = $shareData[0].Server
            $sharePath = $shareData[0].SharePath
            
            $html += @"
            <details class="group-container" $detailsOpen>
                <summary class="subgroup-header">📂 Share: \\$serverName\$shareName $(if($sharePath){"($sharePath)"})</summary>
                <div class="group-content">
"@

            # Group by AD Group within Share
            $adGroups = $shareData | Group-Object -Property ADGroupName
            
            foreach ($adGroup in $adGroups) {
                $groupName = $adGroup.Name
                $groupData = $adGroup.Group
                $rights = $groupData[0].Rights
                $rightsClass = Get-RightsClass -Rights $rights
                
                $html += @"
                <details $detailsOpen>
                    <summary class="subgroup-header">👥 AD Group: $groupName <span style="float:right;padding:4px 10px;border-radius:4px;" class="$rightsClass">$rights</span></summary>
                    <table>
                        <thead>
                            <tr>
                                <th>Domain</th>
                                <th>User</th>
                                <th>Display Name</th>
                                <th>User Group</th>
                                <th>Rights</th>
                            </tr>
                        </thead>
                        <tbody>
"@
                
                foreach ($user in $groupData) {
                    $rightsClass = Get-RightsClass -Rights $user.Rights
                    $html += @"
                            <tr>
                                <td>$($user.Domain)</td>
                                <td>$($user.User)</td>
                                <td>$($user.DisplayName)</td>
                                <td>$($user.UserGroup)</td>
                                <td class="$rightsClass">$($user.Rights)</td>
                            </tr>
"@
                }
                
                $html += @"
                        </tbody>
                    </table>
                </details>
"@
            }
            
            $html += @"
                </div>
            </details>
"@
        }
        
        $html += @"
        </div>
    </details>
"@
    }

    return $html
}

function Generate-PerServerHtml {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [PSCustomObject[]]$Data,

        [Parameter(Mandatory = $false)]
        [switch]$Expandable
    )

    $html = ""
    
    # Group by Server
    $serverGroups = $Data | Group-Object -Property Server | Sort-Object Name

    $detailsOpen = if ($Expandable) { "" } else { "open" }

    foreach ($serverGroup in $serverGroups) {
        $serverName = $serverGroup.Name
        $serverData = $serverGroup.Group
        $serverRecordCount = $serverData.Count
        
        $html += @"
    <details class="group-container" $detailsOpen>
        <summary class="group-header">🖥️ Server: $serverName ($serverRecordCount records)</summary>
        <div class="group-content">
"@

        # Group by Share within Server
        $shareGroups = $serverData | Group-Object -Property Share
        
        foreach ($shareGroup in $shareGroups) {
            $shareName = $shareGroup.Name
            $shareData = $shareGroup.Group
            $sharePath = $shareData[0].SharePath
            $owner1 = $shareData[0].Owner1
            $owner2 = $shareData[0].Owner2
            
            $ownerInfo = ""
            if ($owner1) { $ownerInfo += "Owner: $owner1" }
            if ($owner2) { $ownerInfo += " | Co-Owner: $owner2" }
            
            $html += @"
            <details class="group-container" $detailsOpen>
                <summary class="subgroup-header">📂 Share: $shareName $(if($sharePath){"($sharePath)"}) $(if($ownerInfo){"- $ownerInfo"})</summary>
                <div class="group-content">
"@

            # Group by AD Group within Share
            $adGroups = $shareData | Group-Object -Property ADGroupName
            
            foreach ($adGroup in $adGroups) {
                $groupName = $adGroup.Name
                $groupData = $adGroup.Group
                $rights = $groupData[0].Rights
                $rightsClass = Get-RightsClass -Rights $rights
                
                $html += @"
                <details $detailsOpen>
                    <summary class="subgroup-header">👥 AD Group: $groupName <span style="float:right;padding:4px 10px;border-radius:4px;" class="$rightsClass">$rights</span></summary>
                    <table>
                        <thead>
                            <tr>
                                <th>Domain</th>
                                <th>User</th>
                                <th>Display Name</th>
                                <th>User Group</th>
                                <th>Rights</th>
                            </tr>
                        </thead>
                        <tbody>
"@
                
                foreach ($user in $groupData) {
                    $rightsClass = Get-RightsClass -Rights $user.Rights
                    $html += @"
                            <tr>
                                <td>$($user.Domain)</td>
                                <td>$($user.User)</td>
                                <td>$($user.DisplayName)</td>
                                <td>$($user.UserGroup)</td>
                                <td class="$rightsClass">$($user.Rights)</td>
                            </tr>
"@
                }
                
                $html += @"
                        </tbody>
                    </table>
                </details>
"@
            }
            
            $html += @"
                </div>
            </details>
"@
        }
        
        $html += @"
        </div>
    </details>
"@
    }

    return $html
}

function Generate-XlsxReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [PSCustomObject[]]$Data,

        [Parameter(Mandatory = $true)]
        [string]$ReportType,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath,

        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Summary
    )

    try {
        Import-Module ImportExcel -ErrorAction Stop

        # Create Summary worksheet
        $summaryData = @(
            [PSCustomObject]@{ Metric = "Total Records"; Value = $Summary.TotalRecords }
            [PSCustomObject]@{ Metric = "Unique Servers"; Value = $Summary.UniqueServers }
            [PSCustomObject]@{ Metric = "Unique Shares"; Value = $Summary.UniqueShares }
            [PSCustomObject]@{ Metric = "Unique Owners"; Value = $Summary.UniqueOwners }
            [PSCustomObject]@{ Metric = "Unique Users"; Value = $Summary.UniqueUsers }
            [PSCustomObject]@{ Metric = "Unique AD Groups"; Value = $Summary.UniqueADGroups }
            [PSCustomObject]@{ Metric = "High Risk Permissions"; Value = $Summary.HighRiskCount }
        )

        $summaryData | Export-Excel -Path $OutputPath -WorksheetName "Summary" -AutoSize -BoldTopRow `
            -FreezeTopRow -TableStyle Medium2

        if ($Data.Count -gt 0) {
            # Export full data
            $exportData = $Data | Select-Object Server, Share, SharePath, Owner1, Owner2, ADGroupName, `
                Domain, User, DisplayName, UserGroup, Rights

            $exportData | Export-Excel -Path $OutputPath -WorksheetName "Full Data" -AutoSize -BoldTopRow `
                -FreezeTopRow -TableStyle Medium2

            # Group data based on report type and create separate worksheets
            if ($ReportType -eq "PerServer") {
                $servers = $Data | Select-Object -Property Server -Unique | Sort-Object Server
                
                foreach ($server in $servers) {
                    $serverData = $Data | Where-Object { $_.Server -eq $server.Server }
                    $worksheetName = $server.Server -replace '[\\/:*?"<>|]', '_'
                    # Ensure worksheet name is not empty after sanitization
                    if ([string]::IsNullOrWhiteSpace($worksheetName) -or $worksheetName -match '^_+$') {
                        $worksheetName = "Server_$($servers.IndexOf($server) + 1)"
                    }
                    $worksheetName = $worksheetName.Substring(0, [Math]::Min(31, $worksheetName.Length))
                    
                    $serverData | Select-Object Share, SharePath, Owner1, Owner2, ADGroupName, `
                        Domain, User, DisplayName, UserGroup, Rights | 
                        Export-Excel -Path $OutputPath -WorksheetName $worksheetName -AutoSize `
                        -BoldTopRow -FreezeTopRow -TableStyle Medium2
                }
            }
            elseif ($ReportType -eq "PerOwner") {
                # Get unique owners efficiently
                $owners = @($Data | ForEach-Object { $_.Owner1; $_.Owner2 } | Where-Object { $_ } | Select-Object -Unique | Sort-Object)
                
                $ownerIndex = 0
                foreach ($owner in $owners) {
                    $ownerIndex++
                    $ownerData = $Data | Where-Object { $_.Owner1 -eq $owner -or $_.Owner2 -eq $owner }
                    $worksheetName = $owner -replace '[\\/:*?"<>|]', '_'
                    # Ensure worksheet name is not empty after sanitization
                    if ([string]::IsNullOrWhiteSpace($worksheetName) -or $worksheetName -match '^_+$') {
                        $worksheetName = "Owner_$ownerIndex"
                    }
                    $worksheetName = $worksheetName.Substring(0, [Math]::Min(31, $worksheetName.Length))
                    
                    $ownerData | Select-Object Server, Share, SharePath, ADGroupName, `
                        Domain, User, DisplayName, UserGroup, Rights | 
                        Export-Excel -Path $OutputPath -WorksheetName $worksheetName -AutoSize `
                        -BoldTopRow -FreezeTopRow -TableStyle Medium2
                }
            }

            # Add conditional formatting for high-risk permissions
            $excel = Open-ExcelPackage -Path $OutputPath
            $worksheet = $excel.Workbook.Worksheets["Full Data"]
            
            if ($worksheet) {
                # Find the Rights column
                $rightsCol = $null
                for ($col = 1; $col -le $worksheet.Dimension.End.Column; $col++) {
                    if ($worksheet.Cells[1, $col].Value -eq "Rights") {
                        $rightsCol = $col
                        break
                    }
                }
                
                if ($rightsCol) {
                    # Add conditional formatting
                    $lastRow = $worksheet.Dimension.End.Row
                    $rangeAddress = $worksheet.Cells[2, $rightsCol, $lastRow, $rightsCol].Address
                    
                    # High risk - Red
                    $rule1 = $worksheet.ConditionalFormatting.AddContainsText($rangeAddress)
                    $rule1.Text = "FullControl"
                    $rule1.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $rule1.Style.Fill.BackgroundColor.Color = [System.Drawing.Color]::FromArgb(255, 199, 206)
                    
                    # Medium risk - Orange
                    $rule2 = $worksheet.ConditionalFormatting.AddContainsText($rangeAddress)
                    $rule2.Text = "Modify"
                    $rule2.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $rule2.Style.Fill.BackgroundColor.Color = [System.Drawing.Color]::FromArgb(255, 235, 156)
                }
            }
            
            Close-ExcelPackage $excel
        }

        Write-Verbose "Excel report generated successfully"
    }
    catch {
        Write-Warning "Failed to generate Excel report: $($_.Exception.Message)"
    }
}

function Generate-PdfReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$HtmlPath,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath,

        [Parameter(Mandatory = $true)]
        [string]$ReportType
    )

    try {
        Import-Module PSWritePDF -ErrorAction Stop

        # Read HTML content
        $htmlContent = Get-Content -Path $HtmlPath -Raw -Encoding UTF8

        # Create PDF with header and footer
        $pdfParams = @{
            FilePath        = $OutputPath
            HTML            = $htmlContent
            MarginTop       = 20
            MarginBottom    = 20
            MarginLeft      = 20
            MarginRight     = 20
            PageSize        = 'A4'
            ShowHTML        = $false
        }

        # Generate PDF
        New-PDF @pdfParams

        Write-Verbose "PDF report generated successfully"
    }
    catch {
        Write-Warning "Failed to generate PDF report: $($_.Exception.Message)"
        Write-Verbose "You can manually convert the HTML report to PDF using a web browser or PDF converter."
    }
}

#endregion Helper Functions

# Note: To use as a module, create a .psm1 file and add:
# Export-ModuleMember -Function Generate-ShareAccessReport
