function Get-ComputerSystemInfo {
    <#
    .SYNOPSIS
        Zbiera szczegółowe informacje o komputerze.
    
    .DESCRIPTION
        Funkcja zbiera dane o komputerze obejmujące:
        - Informacje o systemie operacyjnym
        - Informacje o procesorze
        - Informacje o pamięci RAM
        - Informacje o dyskach
        - Informacje o kartach sieciowych
    
    .EXAMPLE
        $computerData = Get-ComputerSystemInfo
        
    .OUTPUTS
        PSCustomObject zawierający wszystkie zebrane informacje o komputerze
    #>
    
    [CmdletBinding()]
    param()
    
    Write-Verbose "Rozpoczynam zbieranie informacji o komputerze..."
    
    # Informacje o systemie operacyjnym
    Write-Verbose "Pobieram informacje o systemie operacyjnym..."
    $os = Get-CimInstance -ClassName Win32_OperatingSystem
    $cs = Get-CimInstance -ClassName Win32_ComputerSystem
    
    $systemInfo = [PSCustomObject]@{
        ComputerName = $cs.Name
        Domain = if ($cs.PartOfDomain) { $cs.Domain } else { "WORKGROUP" }
        Manufacturer = $cs.Manufacturer
        Model = $cs.Model
        OSName = $os.Caption
        OSVersion = $os.Version
        OSArchitecture = $os.OSArchitecture
        InstallDate = $os.InstallDate
        LastBootTime = $os.LastBootUpTime
        CurrentUser = $env:USERNAME
    }
    
    # Informacje o procesorze
    Write-Verbose "Pobieram informacje o procesorze..."
    $processors = Get-CimInstance -ClassName Win32_Processor | ForEach-Object {
        [PSCustomObject]@{
            Name = $_.Name.Trim()
            Manufacturer = $_.Manufacturer
            Cores = $_.NumberOfCores
            LogicalProcessors = $_.NumberOfLogicalProcessors
            MaxClockSpeed = "$($_.MaxClockSpeed) MHz"
            CurrentClockSpeed = "$($_.CurrentClockSpeed) MHz"
            Architecture = switch ($_.Architecture) {
                0 { "x86" }
                1 { "MIPS" }
                2 { "Alpha" }
                3 { "PowerPC" }
                5 { "ARM" }
                6 { "Itanium" }
                9 { "x64" }
                default { "Unknown" }
            }
        }
    }
    
    # Informacje o pamięci RAM
    Write-Verbose "Pobieram informacje o pamięci RAM..."
    $totalRAM = [math]::Round($cs.TotalPhysicalMemory / 1GB, 2)
    $freeRAM = [math]::Round($os.FreePhysicalMemory / 1MB, 2)
    $usedRAM = [math]::Round(($cs.TotalPhysicalMemory - ($os.FreePhysicalMemory * 1KB)) / 1GB, 2)
    
    $memoryInfo = [PSCustomObject]@{
        TotalRAM_GB = $totalRAM
        UsedRAM_GB = $usedRAM
        FreeRAM_MB = $freeRAM
        UsagePercent = [math]::Round(($usedRAM / $totalRAM) * 100, 2)
    }
    
    # Szczegóły modułów pamięci
    $memoryModules = Get-CimInstance -ClassName Win32_PhysicalMemory | ForEach-Object {
        [PSCustomObject]@{
            BankLabel = $_.BankLabel
            Capacity_GB = [math]::Round($_.Capacity / 1GB, 2)
            Speed = "$($_.Speed) MHz"
            Manufacturer = $_.Manufacturer
            PartNumber = $_.PartNumber.Trim()
        }
    }
    
    # Informacje o dyskach
    Write-Verbose "Pobieram informacje o dyskach..."
    $disks = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3" | ForEach-Object {
        $size = [math]::Round($_.Size / 1GB, 2)
        $free = [math]::Round($_.FreeSpace / 1GB, 2)
        $used = $size - $free
        
        [PSCustomObject]@{
            Drive = $_.DeviceID
            VolumeName = $_.VolumeName
            FileSystem = $_.FileSystem
            TotalSize_GB = $size
            UsedSpace_GB = $used
            FreeSpace_GB = $free
            UsagePercent = if ($size -gt 0) { [math]::Round(($used / $size) * 100, 2) } else { 0 }
        }
    }
    
    # Informacje o dyskach fizycznych
    $physicalDisks = Get-CimInstance -ClassName Win32_DiskDrive | ForEach-Object {
        [PSCustomObject]@{
            Model = $_.Model
            InterfaceType = $_.InterfaceType
            Size_GB = [math]::Round($_.Size / 1GB, 2)
            Partitions = $_.Partitions
            Status = $_.Status
        }
    }
    
    # Informacje o kartach sieciowych
    Write-Verbose "Pobieram informacje o kartach sieciowych..."
    $networkAdapters = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter "IPEnabled=True" | ForEach-Object {
        [PSCustomObject]@{
            Description = $_.Description
            MACAddress = $_.MACAddress
            IPAddress = $_.IPAddress -join ", "
            SubnetMask = $_.IPSubnet -join ", "
            DefaultGateway = $_.DefaultIPGateway -join ", "
            DNSServers = $_.DNSServerSearchOrder -join ", "
            DHCPEnabled = $_.DHCPEnabled
        }
    }
    
    # Informacje o BIOS
    Write-Verbose "Pobieram informacje o BIOS..."
    $bios = Get-CimInstance -ClassName Win32_BIOS
    $biosInfo = [PSCustomObject]@{
        Manufacturer = $bios.Manufacturer
        Version = $bios.SMBIOSBIOSVersion
        ReleaseDate = $bios.ReleaseDate
        SerialNumber = $bios.SerialNumber
    }
    
    # Zwrócenie wszystkich zebranych danych
    Write-Verbose "Zakończono zbieranie informacji."
    
    return [PSCustomObject]@{
        CollectionDate = Get-Date
        SystemInfo = $systemInfo
        BIOSInfo = $biosInfo
        Processors = $processors
        Memory = $memoryInfo
        MemoryModules = $memoryModules
        LogicalDisks = $disks
        PhysicalDisks = $physicalDisks
        NetworkAdapters = $networkAdapters
    }
}

# Jeśli skrypt jest uruchamiany bezpośrednio (nie importowany jako moduł)
if ($MyInvocation.InvocationName -ne '.') {
    Get-ComputerSystemInfo
}
