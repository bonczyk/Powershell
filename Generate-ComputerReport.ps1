# ====================================================================
# Główny skrypt do generowania raportu o komputerze
# ====================================================================
# 
# Ten skrypt łączy dwie procedury:
# 1. Get-ComputerSystemInfo - zbiera dane o komputerze
# 2. New-ComputerHTMLReport - generuje raport HTML
#
# Użycie:
#   .\Generate-ComputerReport.ps1
#   .\Generate-ComputerReport.ps1 -OutputPath "C:\Reports\MyReport.html"
#   .\Generate-ComputerReport.ps1 -OpenReport
#   .\Generate-ComputerReport.ps1 -Verbose
#
# ====================================================================

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$OutputPath,
    
    [Parameter(Mandatory=$false)]
    [switch]$OpenReport
)

# Pobierz ścieżkę do katalogu ze skryptami
$ScriptPath = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path $MyInvocation.MyCommand.Path }

Write-Host "`n╔════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║        Generator Raportu Informacji o Komputerze          ║" -ForegroundColor Cyan
Write-Host "╚════════════════════════════════════════════════════════════╝`n" -ForegroundColor Cyan

try {
    # Krok 1: Zbierz informacje o komputerze
    Write-Host "[1/2] Zbieranie informacji o komputerze..." -ForegroundColor Yellow
    
    # Importuj i uruchom funkcję zbierającą dane
    . "$ScriptPath\Get-ComputerInfo.ps1"
    $computerData = Get-ComputerSystemInfo -Verbose:$VerbosePreference
    
    Write-Host "      ✓ Zebrano dane o komputerze" -ForegroundColor Green
    Write-Host "        - System: $($computerData.SystemInfo.OSName)" -ForegroundColor Gray
    Write-Host "        - Komputer: $($computerData.SystemInfo.ComputerName)" -ForegroundColor Gray
    Write-Host "        - Procesory: $($computerData.Processors.Count)" -ForegroundColor Gray
    Write-Host "        - RAM: $($computerData.Memory.TotalRAM_GB) GB" -ForegroundColor Gray
    Write-Host "        - Dyski: $($computerData.LogicalDisks.Count) logiczne, $($computerData.PhysicalDisks.Count) fizyczne`n" -ForegroundColor Gray
    
    # Krok 2: Wygeneruj raport HTML
    Write-Host "[2/2] Generowanie raportu HTML..." -ForegroundColor Yellow
    
    # Importuj i uruchom funkcję generującą raport
    . "$ScriptPath\New-ComputerReport.ps1"
    
    $reportParams = @{
        ComputerData = $computerData
        OpenReport = $OpenReport
        Verbose = $VerbosePreference
    }
    
    if ($OutputPath) {
        $reportParams.OutputPath = $OutputPath
    }
    
    $reportPath = New-ComputerHTMLReport @reportParams
    
    Write-Host "`n╔════════════════════════════════════════════════════════════╗" -ForegroundColor Green
    Write-Host "║                   SUKCES!                                  ║" -ForegroundColor Green
    Write-Host "╚════════════════════════════════════════════════════════════╝" -ForegroundColor Green
    Write-Host "`nRaport został pomyślnie wygenerowany!" -ForegroundColor Green
    Write-Host "Lokalizacja: $reportPath`n" -ForegroundColor Cyan
    
    if (-not $OpenReport) {
        Write-Host "Aby otworzyć raport w przeglądarce, użyj:" -ForegroundColor Yellow
        Write-Host "  Invoke-Item `"$reportPath`"`n" -ForegroundColor Gray
    }
}
catch {
    Write-Host "`n╔════════════════════════════════════════════════════════════╗" -ForegroundColor Red
    Write-Host "║                    BŁĄD!                                   ║" -ForegroundColor Red
    Write-Host "╚════════════════════════════════════════════════════════════╝" -ForegroundColor Red
    Write-Host "`nWystąpił błąd podczas generowania raportu:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host "`nSzczegóły:" -ForegroundColor Yellow
    Write-Host $_.ScriptStackTrace -ForegroundColor Gray
    exit 1
}
