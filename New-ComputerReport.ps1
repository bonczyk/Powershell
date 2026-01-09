function New-ComputerHTMLReport {
    <#
    .SYNOPSIS
        Generuje profesjonalny raport HTML z informacjami o komputerze.
    
    .DESCRIPTION
        Funkcja przyjmuje dane zebrane przez Get-ComputerSystemInfo i generuje
        estetyczny, profesjonalny raport w formacie HTML z tabelami i stylami CSS.
    
    .PARAMETER ComputerData
        Obiekt zawierający dane o komputerze (output z Get-ComputerSystemInfo)
    
    .PARAMETER OutputPath
        Ścieżka do pliku HTML (domyślnie: ComputerReport_[data].html w bieżącym katalogu)
    
    .PARAMETER OpenReport
        Przełącznik - jeśli podany, raport zostanie automatycznie otwarty w przeglądarce
    
    .EXAMPLE
        $data = Get-ComputerSystemInfo
        New-ComputerHTMLReport -ComputerData $data -OpenReport
        
    .EXAMPLE
        Get-ComputerSystemInfo | New-ComputerHTMLReport -OutputPath "C:\Reports\MyPC.html"
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [PSCustomObject]$ComputerData,
        
        [Parameter(Mandatory=$false)]
        [string]$OutputPath,
        
        [Parameter(Mandatory=$false)]
        [switch]$OpenReport
    )
    
    # Ustal ścieżkę wyjściową
    if (-not $OutputPath) {
        $date = Get-Date -Format "yyyy-MM-dd_HHmmss"
        $OutputPath = Join-Path $PWD "ComputerReport_$date.html"
    }
    
    Write-Verbose "Generuję raport HTML..."
    
    # Funkcja pomocnicza do konwersji obiektu na wiersze tabeli HTML
    function ConvertTo-HTMLTableRows {
        param($Object, $Properties)
        
        $rows = ""
        foreach ($item in $Object) {
            $rows += "<tr>"
            foreach ($prop in $Properties) {
                $value = $item.$prop
                if ($null -eq $value) { $value = "N/A" }
                $rows += "<td>$value</td>"
            }
            $rows += "</tr>"
        }
        return $rows
    }
    
    # Funkcja pomocnicza do konwersji pojedynczego obiektu na tabelę klucz-wartość
    function ConvertTo-HTMLKeyValueTable {
        param($Object)
        
        $rows = ""
        foreach ($prop in $Object.PSObject.Properties) {
            $rows += "<tr><th>$($prop.Name)</th><td>$($prop.Value)</td></tr>"
        }
        return $rows
    }
    
    # CSS Style
    $cssStyle = @"
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 20px;
            color: #333;
        }
        
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            border-radius: 10px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.2);
            overflow: hidden;
        }
        
        .header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 40px;
            text-align: center;
        }
        
        .header h1 {
            font-size: 2.5em;
            margin-bottom: 10px;
            font-weight: 300;
        }
        
        .header .subtitle {
            font-size: 1.1em;
            opacity: 0.9;
        }
        
        .content {
            padding: 40px;
        }
        
        .section {
            margin-bottom: 40px;
        }
        
        .section-title {
            font-size: 1.8em;
            color: #667eea;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 3px solid #667eea;
            font-weight: 500;
        }
        
        table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 20px;
            background: white;
            box-shadow: 0 2px 10px rgba(0,0,0,0.05);
            border-radius: 8px;
            overflow: hidden;
        }
        
        th {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 15px;
            text-align: left;
            font-weight: 600;
            text-transform: uppercase;
            font-size: 0.9em;
            letter-spacing: 0.5px;
        }
        
        td {
            padding: 12px 15px;
            border-bottom: 1px solid #f0f0f0;
        }
        
        tr:hover {
            background-color: #f8f9ff;
        }
        
        tr:last-child td {
            border-bottom: none;
        }
        
        .kv-table th {
            width: 40%;
            background: #f8f9ff;
            color: #667eea;
            font-weight: 600;
        }
        
        .summary-cards {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            margin-bottom: 40px;
        }
        
        .card {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 25px;
            border-radius: 10px;
            box-shadow: 0 4px 15px rgba(102, 126, 234, 0.3);
        }
        
        .card-title {
            font-size: 0.9em;
            opacity: 0.9;
            margin-bottom: 10px;
            text-transform: uppercase;
            letter-spacing: 1px;
        }
        
        .card-value {
            font-size: 2em;
            font-weight: 600;
        }
        
        .footer {
            text-align: center;
            padding: 20px;
            background: #f8f9ff;
            color: #666;
            font-size: 0.9em;
        }
        
        .warning {
            background-color: #fff3cd;
            color: #856404;
            padding: 10px;
            border-left: 4px solid #ffc107;
        }
        
        .success {
            background-color: #d4edda;
            color: #155724;
            padding: 10px;
            border-left: 4px solid #28a745;
        }
    </style>
"@
    
    # Przygotowanie danych dla kart podsumowania
    $totalDiskSpace = ($ComputerData.LogicalDisks | Measure-Object -Property TotalSize_GB -Sum).Sum
    $usedDiskSpace = ($ComputerData.LogicalDisks | Measure-Object -Property UsedSpace_GB -Sum).Sum
    $diskUsagePercent = if ($totalDiskSpace -gt 0) { [math]::Round(($usedDiskSpace / $totalDiskSpace) * 100, 1) } else { 0 }
    
    # Generowanie HTML
    $html = @"
<!DOCTYPE html>
<html lang="pl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Raport Systemowy - $($ComputerData.SystemInfo.ComputerName)</title>
    $cssStyle
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🖥️ Raport Systemowy</h1>
            <div class="subtitle">$($ComputerData.SystemInfo.ComputerName) - $($ComputerData.SystemInfo.OSName)</div>
            <div class="subtitle">Wygenerowano: $($ComputerData.CollectionDate.ToString("yyyy-MM-dd HH:mm:ss"))</div>
        </div>
        
        <div class="content">
            <!-- Karty podsumowania -->
            <div class="summary-cards">
                <div class="card">
                    <div class="card-title">Procesory</div>
                    <div class="card-value">$($ComputerData.Processors.Count)</div>
                </div>
                <div class="card">
                    <div class="card-title">Pamięć RAM</div>
                    <div class="card-value">$($ComputerData.Memory.TotalRAM_GB) GB</div>
                </div>
                <div class="card">
                    <div class="card-title">Wykorzystanie RAM</div>
                    <div class="card-value">$($ComputerData.Memory.UsagePercent)%</div>
                </div>
                <div class="card">
                    <div class="card-title">Dyski logiczne</div>
                    <div class="card-value">$($ComputerData.LogicalDisks.Count)</div>
                </div>
            </div>
            
            <!-- Informacje o systemie -->
            <div class="section">
                <h2 class="section-title">📋 Informacje o systemie</h2>
                <table class="kv-table">
                    $(ConvertTo-HTMLKeyValueTable $ComputerData.SystemInfo)
                </table>
            </div>
            
            <!-- Informacje o BIOS -->
            <div class="section">
                <h2 class="section-title">⚙️ Informacje o BIOS</h2>
                <table class="kv-table">
                    $(ConvertTo-HTMLKeyValueTable $ComputerData.BIOSInfo)
                </table>
            </div>
            
            <!-- Procesory -->
            <div class="section">
                <h2 class="section-title">🔧 Procesory</h2>
                <table>
                    <thead>
                        <tr>
                            <th>Nazwa</th>
                            <th>Producent</th>
                            <th>Rdzenie</th>
                            <th>Wątki logiczne</th>
                            <th>Max. częstotliwość</th>
                            <th>Architektura</th>
                        </tr>
                    </thead>
                    <tbody>
                        $(ConvertTo-HTMLTableRows $ComputerData.Processors @('Name', 'Manufacturer', 'Cores', 'LogicalProcessors', 'MaxClockSpeed', 'Architecture'))
                    </tbody>
                </table>
            </div>
            
            <!-- Pamięć RAM -->
            <div class="section">
                <h2 class="section-title">💾 Pamięć RAM</h2>
                <table class="kv-table">
                    $(ConvertTo-HTMLKeyValueTable $ComputerData.Memory)
                </table>
                
                <h3 style="margin-top: 30px; margin-bottom: 15px; color: #667eea;">Moduły pamięci</h3>
                <table>
                    <thead>
                        <tr>
                            <th>Bank</th>
                            <th>Pojemność (GB)</th>
                            <th>Częstotliwość</th>
                            <th>Producent</th>
                            <th>Nr katalogowy</th>
                        </tr>
                    </thead>
                    <tbody>
                        $(ConvertTo-HTMLTableRows $ComputerData.MemoryModules @('BankLabel', 'Capacity_GB', 'Speed', 'Manufacturer', 'PartNumber'))
                    </tbody>
                </table>
            </div>
            
            <!-- Dyski logiczne -->
            <div class="section">
                <h2 class="section-title">💿 Dyski logiczne</h2>
                <table>
                    <thead>
                        <tr>
                            <th>Dysk</th>
                            <th>Nazwa woluminu</th>
                            <th>System plików</th>
                            <th>Rozmiar (GB)</th>
                            <th>Wykorzystane (GB)</th>
                            <th>Wolne (GB)</th>
                            <th>Wykorzystanie (%)</th>
                        </tr>
                    </thead>
                    <tbody>
                        $(ConvertTo-HTMLTableRows $ComputerData.LogicalDisks @('Drive', 'VolumeName', 'FileSystem', 'TotalSize_GB', 'UsedSpace_GB', 'FreeSpace_GB', 'UsagePercent'))
                    </tbody>
                </table>
            </div>
            
            <!-- Dyski fizyczne -->
            <div class="section">
                <h2 class="section-title">🗄️ Dyski fizyczne</h2>
                <table>
                    <thead>
                        <tr>
                            <th>Model</th>
                            <th>Interfejs</th>
                            <th>Rozmiar (GB)</th>
                            <th>Partycje</th>
                            <th>Status</th>
                        </tr>
                    </thead>
                    <tbody>
                        $(ConvertTo-HTMLTableRows $ComputerData.PhysicalDisks @('Model', 'InterfaceType', 'Size_GB', 'Partitions', 'Status'))
                    </tbody>
                </table>
            </div>
            
            <!-- Karty sieciowe -->
            <div class="section">
                <h2 class="section-title">🌐 Karty sieciowe</h2>
                <table>
                    <thead>
                        <tr>
                            <th>Opis</th>
                            <th>Adres MAC</th>
                            <th>Adres IP</th>
                            <th>Maska podsieci</th>
                            <th>Brama domyślna</th>
                            <th>Serwery DNS</th>
                            <th>DHCP</th>
                        </tr>
                    </thead>
                    <tbody>
                        $(ConvertTo-HTMLTableRows $ComputerData.NetworkAdapters @('Description', 'MACAddress', 'IPAddress', 'SubnetMask', 'DefaultGateway', 'DNSServers', 'DHCPEnabled'))
                    </tbody>
                </table>
            </div>
        </div>
        
        <div class="footer">
            Raport wygenerowany automatycznie przez PowerShell<br>
            © $(Get-Date -Format yyyy) - Raport komputera: $($ComputerData.SystemInfo.ComputerName)
        </div>
    </div>
</body>
</html>
"@
    
    # Zapisanie raportu do pliku
    try {
        $html | Out-File -FilePath $OutputPath -Encoding UTF8 -Force
        Write-Host "✓ Raport został wygenerowany: $OutputPath" -ForegroundColor Green
        
        # Otwarcie raportu w przeglądarce jeśli podano przełącznik
        if ($OpenReport) {
            Write-Verbose "Otwieram raport w przeglądarce..."
            Invoke-Item $OutputPath
        }
        
        # Zwrócenie ścieżki do wygenerowanego raportu
        return $OutputPath
    }
    catch {
        Write-Error "Błąd podczas generowania raportu: $_"
        throw
    }
}

# Jeśli skrypt jest uruchamiany bezpośrednio
if ($MyInvocation.InvocationName -ne '.') {
    Write-Host "`nTen skrypt wymaga danych z Get-ComputerSystemInfo." -ForegroundColor Yellow
    Write-Host "Przykładowe użycie:" -ForegroundColor Cyan
    Write-Host "  `$data = .\Get-ComputerInfo.ps1" -ForegroundColor Gray
    Write-Host "  .\New-ComputerReport.ps1 -ComputerData `$data -OpenReport" -ForegroundColor Gray
    Write-Host "`nlub:" -ForegroundColor Cyan
    Write-Host "  .\Get-ComputerInfo.ps1 | .\New-ComputerReport.ps1 -OpenReport`n" -ForegroundColor Gray
}
