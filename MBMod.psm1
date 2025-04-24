<#
 Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1"; cd DUB:
 Import-Module "\\drsitsrv1\DRSsupport$\Projects\2022\Test-BCS\modules\MBMod\0.3\MBMod.psm1" -Force -WarningAction SilentlyContinue
 Import-Module "H:\MB\PS\modules\MBMod\0.3\MBMod.psm1" -Force -WarningAction SilentlyContinue
 Import-Module "\\fxt8\c$\H\MB\PS\modules\MBMod\0.3\MBMod.psm1" -Force -WarningAction SilentlyContinue
 [Management.Automation.WildcardPattern]::Escape('test[1].txt (foo)')
 [regex]::Escape("test[1].txt (foo)")
 Run-Remote $pc "powershell -command `"Start-Transcript c:\temp\log_appx.txt; Get-appxprovisionedpackage –online -Verbose | where-object {`$_.displayname -like \`"*Edge*\`" }; Stop-Transcript`""
#> 
<# adinfo
 $l = Ping-DealersPCs

 $out = $l | % {  [PSCustomObject]@{ PC = $_; folder=(Test-path "\\$_\c$\Program Files\DRS" ) } }
 $a = $out.folder | % { [PSCustomObject]@{ PC = $_.DirectoryName; size=$_.Length } }

 $list = $l | % { $path ="\\$_\c$\Temp\Logs\11023_*.txt";  if (Test-Path $path) { [PSCustomObject]@{ PC = $_; folder=(gci $path ) } } }
 $list.folder | % { PraseNetUse (gc $_) } | select -Unique Remote
 
$MemoryStream = [System.IO.MemoryStream]::new()
$Compressor = [System.IO.Compression.DeflateStream]::new($MemoryStream,[System.IO.Compression.CompressionMode]::Compress)
$CompressionWriter = [System.IO.StreamWriter]::new($Compressor)
$CompressionWriter.Write($mystring)
$CompressedByteArray = $MemoryStream.ToArray()

#>

function CRQ-Edge {

  $date = Get-Date
  $a = 'while ($Date.DayOfWeek -notin "Tuesday","Thursday") {$date = $date.AddDays(1);}'
  $ver = (Get-CMApplication -Fast -Name *Edge*).SoftwareVersion | sort | select -Last 1

  @"
Microsoft Edge $ver

Brief description of Change?

The change involves the rollout of a new version of Microsoft Edge $ver to all company workstations. This update aims to enhance security, performance, and introduce new features.

When is it scheduled to be implemented? (Start & End Dates/Times)

$(Get-Date $date -f d) – Out to test PC
$(Get-Date $date -f d) – Out to test Group
$($date=$date.AddDays(6);iex $a;Get-Date $date -f d) – Out to Group 1 – approximately 25% of Dealers PC’s
$($date=$date.AddDays(1);iex $a;Get-Date $date -f d) – Out to Group 2 – approximately 25% of Dealers PC’s
$($date=$date.AddDays(1);iex $a;Get-Date $date -f d) - Out to Group 3 – approximately 25% of Dealers PC’s
$($date=$date.AddDays(1);iex $a;Get-Date $date -f d) - Out to Group 4 – approximately 25% of Dealers PC’s


In the worst-case scenario, what services could be impacted?

In the worst-case scenario, users might experience temporary disruption in accessing the web browser, which could impact web-based applications and services dependent on Edge.


Has support from the required Teams for implementing/testing this Change been confirmed?

Yes, support from the Dealing Room Support team, ready to assist during the implementation and testing phases.


Are you aware of any possible impacts from this Change being implemented in the same time frame as other changes?

There are no known conflicts with other scheduled changes during this time frame. Coordination has been done to ensure no overlap with other major updates or network maintenance activities.


What validation (production testing post-deployment) will be carried out?

The application package has already been tested on the test computer and user testing has been performed in Molesworth.


Has the back-out plan in place been tested?

Yes, the back-out plan includes reverting to the previous stable version of Microsoft Edge and ensuring all user data and settings are preserved.
"@
}

function CRQ-CU {

  $date = Get-Date
  $a = 'while ($Date.DayOfWeek -notin "Tuesday","Thursday") {$date = $date.AddDays(1);}'

  @"
$(Get-Date -f yyyy-MM) Monthly Updates required to maintain integrity and security of dealers desktops. 	


Brief description of Change?

The change involves the rollout of a new Microsoft Windows updatetes and patches to all dealers workstations. This update aims to enhance security, performance, and introduce new features.

$($global:kbs.title -join "`n")                           


When is it scheduled to be implemented? (Start & End Dates/Times)

$(Get-Date $date -f d) – Out to test PC
$(Get-Date $date -f d) – Out to test Group
$($date=$date.AddDays(6);iex $a;Get-Date $date -f d) – Out to Group 1 – approximately 25% of Dealers PC’s
$($date=$date.AddDays(1);iex $a;Get-Date $date -f d) – Out to Group 2 – approximately 25% of Dealers PC’s
$($date=$date.AddDays(1);iex $a;Get-Date $date -f d) - Out to Group 3 – approximately 25% of Dealers PC’s
$($date=$date.AddDays(1);iex $a;Get-Date $date -f d) - Out to Group 4 – approximately 25% of Dealers PC’s


In the worst-case scenario, what services could be impacted?

In the worst-case scenario, users might experience issues with Windows 10


Has support from the required Teams for implementing/testing this Change been confirmed?

Yes, support from the Dealing Room Support team, ready to assist during the implementation and testing phases.


Are you aware of any possible impacts from this Change being implemented in the same time frame as other changes?

There are no known conflicts with other scheduled changes during this time frame. Coordination has been done to ensure no overlap with other major updates or network maintenance activities.


What validation (production testing post-deployment) will be carried out?

The application package has already been tested on the test computer and user testing has been performed in Molesworth.


Has the back-out plan in place been tested?

Yes, the back-out plan includes reverting to the previous state, uninstalling updates and ensuring all user data and settings are preserved.
"@
}

function Pack-CU {
  Save-NewUpdate
  ExtractCabsFolder
  Move-toCM
  New-MSPapp
  New-MSUapp
  DeployToGroup
}

function SCCM-NewApp {
  param (
    [Parameter(Mandatory)][string]$AppName,
    [Parameter(Mandatory)][string]$Description,
    [Parameter(Mandatory)][string]$Publisher,
    [Parameter(Mandatory)][string]$SoftwareVersion,
    [Parameter(Mandatory)][string]$Icon,
    [Parameter(Mandatory)][string]$ContentLocation,
    [Parameter(Mandatory)][string]$InstallCommand,
    [string]$DTName = "", 
    [string]$FolderPath = "",
    [string]$DPName = "",
    [string]$DPGroupName = "AllDP",
    [int]$EstimatedRuntimeMins = 10,
    [PSCustomObject]$DetectionClause
  )
  Import-Module (Join-Path $(Split-Path $env:SMS_ADMIN_UI_PATH) ConfigurationManager.psd1)
  IF ($(pwd).path -ne "$SCCMSiteCode`:\") { Set-Location "$SCCMSiteCode`:" }
  $newApp = @{ Name = $AppName; Description = $Description; Publisher = $Publisher; SoftwareVersion = $SoftwareVersion; IconLocationFile = $Icon }
  $newApp | ft
  $app = Get-CMApplication -Fast -Name $AppName
  if (!($app)) { $app = New-CMApplication @newApp }
  $app | Select LocalizedDisplayName, LocalizedDescription
  $addDT = @{
    ApplicationName          = $AppName
    DeploymentTypeName       = $DTName
    InstallCommand           = $InstallCommand
    ContentLocation          = $ContentLocation
    InstallationBehaviorType = 'InstallForSystem'
    EstimatedRuntimeMins     = $EstimatedRuntimeMins
    MaximumRuntimeMins       = 15
    LogonRequirementType     = 'WhetherOrNotUserLoggedOn'
    ScriptLanguage           = 'PowerShell'
    ScriptText               = ''
    Comment                  = "$(get-date) - $AppName"
  }
  $addDT | ft 
  if ($DetectionClause) { Add-CMScriptDeploymentType @addDT -AddDetectionClause $DetectionClause | Select LocalizedDisplayName, LocalizedDescription }
  else { 'ScriptLanguage', 'ScriptText' | % { $addDT.Remove($_) }; Add-CMMsiDeploymentType @addDT | Select LocalizedDescription, LocalizedDisplayName }    
  if ($FolderPath) { Move-CMObject -FolderPath $FolderPath -InputObject $app }
  if ($DPName) { Start-CMContentDistribution -InputObject $app -DistributionPointName $DPName -ErrorAction SilentlyContinue }
  elseif ($DPGroupName) { Start-CMContentDistribution -InputObject $app -DistributionPointGroupName $DPGroupName -ErrorAction SilentlyContinue }
}
function Pack-Java {
  param (
    [string]$Location = "\\drscmsrv2\e$\SoftwarePackages\Java\",
    [string]$SCCM = "\\drscmsrv2\e$\SoftwarePackages\Java"
  )
    
  # Locate the latest Java installer
  $Path = (Get-ChildItem -Path $Location -Filter "jre*.exe" -Recurse | Sort-Object LastAccessTime | Select-Object -Last 1).FullName
  if (-not $Path) {
    Write-Error "No Java installer found in $Location"
    return
  }

  $File = Get-Item $Path
  $Info = $File.VersionInfo
  $FileVer = $Info.FileVersion
  $DestinationFolder = Join-Path -Path $SCCM -ChildPath $FileVer

  # Log details
  Write-Output "Downloaded version: $FileVer"
  Write-Output "Destination folder is $DestinationFolder"

  # Ensure the destination folder exists
  if (-not (Test-Path $DestinationFolder)) {
    Write-Output "Creating destination folder: $DestinationFolder"
    New-Item -ItemType Directory -Path $DestinationFolder -Force
    Move-Item -Path $Path -Destination (Join-Path -Path $DestinationFolder -ChildPath $File.Name)
  }

  # Load SCCM Module and Map Drive
  SCCM-MapDrive
  SCCM-LoadModule

  # Check if deployment already exists
  if (-not (Get-CMDeploymentType -ApplicationName "Java" -DeploymentTypeName "Java $FileVer")) {
    Write-Output "No deployment type exists for Java - $FileVer"

    # Use SCCM-NewApp for simplified SCCM application creation
    $NewAppParams = @{
      AppName              = "Java $FileVer"
      Description          = "$($Info.FileDescription) - $FileVer"
      Publisher            = $Info.CompanyName
      SoftwareVersion      = $FileVer
      Icon                 = '\\drscmsrv2\e$\SoftwarePackages\_ico\java_original_logo_icon_146458.png'
      ContentLocation      = $DestinationFolder
      InstallCommand       = "Java.bat"
      DTName               = "DT_Java_$FileVer"
      FolderPath           = "DUB:\Application\Java"
      DPGroupName          = "AllDP"
      EstimatedRuntimeMins = 10
    }
    SCCM-NewApp @NewAppParams

    # Distribute and deploy the application
    Start-CMContentDistribution -ApplicationName "Java $FileVer" -DistributionPointName 'drscmsrv2.dealers.aib.pri' -DistributionPointGroupName 'AllDP'
    DeployToGroup -Apps "Java $FileVer" -Collection "Test_MB" -Now
    Invoke-CMClientNotification -ActionType ClientNotificationRequestMachinePolicyNow -CollectionName "Test_MB"
  }
  else {
    Write-Output "Deployment type for Java $FileVer already exists."
  }
}
function Pack-Java {
  $Location = "\\drscmsrv2\e$\SoftwarePackages\Java\"
  $Path = (gci $Location "jre*.exe" -Recurse | sort LastAccessTime | select -Last 1).fullname
  $Path
  $SCCM = '\\drscmsrv2\e$\SoftwarePackages\Java'
  $file = gci $path
  $info = (gci $path).VersionInfo
  hl "Downloaded version: $EdgeVersion" $EdgeVersion 
  hl "Destination folder: $destinationfolder" "$destinationfolder" 
  SCCM-MapDrive

  $FileVer = $info.FileVersion
  $FileName = $file.Name
  $destinationfolder = "$SCCM\$FileVer"
  Write-Output "Downloaded version: $FileVer"
  Write-Output "Destination folder is $destinationfolder"

  $JavaVer3d = (($FileVer -split '\.')[2]).Substring(0, 3)
  # $javaGUID = "{77924AE4-039E-4CA4-87B4-2F32180$($JavaVer3d)F0}"
  $javaGUID2 = "{71024AE4-039E-4CA4-87B4-2F32180$($JavaVer3d)F0}"
  IF (!(test-path $destinationfolder)) {
    hl "Creating $destinationfolder" "$destinationfolder"
    [System.IO.Directory]::CreateDirectory($destinationfolder); Write-Output "Moving $Path to $destinationfolder"  
    [System.IO.File]::Move($Path, "$destinationfolder\$Filename")  
  }
  SCCM-LoadModule
  IF (!(Get-CMDeploymentType -ApplicationName "Java" -DeploymentTypeName "Java $FileVer")) {
    Write-Output "No deployment type exists for Java - $FileVer"

    $newApp = @{ Name  = "Java $FileVer"
      Description      = "$($info.FileDescription) - $($FileVer) - $($JavaVer3d) - $($javaGUID)"
      Publisher        = $info.CompanyName
      SoftwareVersion  = $FileVer
      IconLocationFile = '\\drscmsrv2\e$\SoftwarePackages\_ico\java_original_logo_icon_146458.png'
    }
    $newApp | ft
    New-CMApplication @newApp | Select LocalizedDescription, LocalizedDisplayName

    $addMsi = @{ ApplicationName = "Java $FileVer"
      DeploymentTypeName         = "DT_Java_$FileVer"
      InstallCommand             = 'Java.bat'
      ContentLocation            = "$destinationfolder"
      InstallationBehaviorType   = 'InstallForSystem' 
      EstimatedRuntimeMins       = 5 
      LogonRequirementType       = 'WhetherOrNotUserLoggedOn'
      ScriptLanguage             = 'PowerShell'
      ScriptText                 = ''
      Comment                    = "$(get-date) - $($FileName) - $($JavaVer3d)"  
    }
    $addMsi | ft
    Add-CMScriptDeploymentType @addMsi | Select LocalizedDescription, LocalizedDisplayName
    $cl1 = New-CMDetectionClauseWindowsInstaller -ProductCode $javaGUID -Existence
    Set-CMScriptDeploymentType -ApplicationName "Java $FileVer" -DeploymentTypeName "DT_Java_$FileVer" -AddDetectionClause $cl1
    $a = Get-CMApplication -Name "Java $FileVer"
    Move-CMObject -FolderPath "DUB:\Application\Java" -InputObject $a

    "Add files to deployment folder"
    cd c: ; ii $Location 
    pause
    Start-CMContentDistribution -ApplicationName "Java $FileVer" -DistributionPointName 'drscmsrv2.dealers.aib.pri' -DistributionPointGroupName 'AllDP'

    $NewDep = @{ ApplicationName = "Java $FileVer"
      CollectionName             = "Test_MB"
      AvailableDateTime          = get-date -Hour 22 -Minute 15
      DeadlineDateTime           = get-date
      DeployAction               = "Install"                
      DeployPurpose              = "Required"
      UserNotification           = "DisplaySoftwareCenterOnly"
      SendWakeupPacket           = $true  
      PersistOnWriteFilterDevice = $false
    }
    $NewDep | ft
    New-CMApplicationDeployment @NewDep | select ApplicationName, CollectionName, StartTime
    Invoke-CMClientNotification -ActionType ClientNotificationRequestMachinePolicyNow -CollectionName "Test_MB"
  }
  ELSE { Write-Output "$destinationfolder already exists" }
  Set-Location $SavedPath
}    

function SCCM-MapDrive {
  param ($RemotePath = '\\drscmsrv2\e$', $UserName = 'adm_58691', $freeletter = ( ls function:[d-z]: -n | ? { !(Test-Path $_ -EA SilentlyContinue) } | select -Last 1) )
  if ((Get-SmbMapping).RemotePath -notcontains $RemotePath) {
    if ($freeletter) {
      $p = Read-Host "Enter Password" -AsSecureString
      New-SmbMapping -LocalPath $freeletter -RemotePath $RemotePath -UserName $UserName -Password ([Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($p)))
    }
    else { Write-Error "No available drive letters." }
  }
  else { Write-Host "Path already mapped." }
}

function SCCM-LoadModule($SCCMSiteCode = 'DUB') {
  Import-Module (Join-Path $(Split-Path $env:SMS_ADMIN_UI_PATH) ConfigurationManager.psd1)
  IF ($(pwd).path -ne "$SCCMSiteCode`:\") { Set-Location "$SCCMSiteCode`:" }
}

function Pack-Edge($Path = "\\drscmsrv2\e$\SoftwarePackages\Microsoft EDGE\", $SCCM = '\\drscmsrv2\e$\SoftwarePackages\Microsoft EDGE') {
  cd C:
  $path = (gci -Path $Path -Filter *.msi -Recurse | sort LastWriteTime | select -Last 1).FullName
  $Meta = Get-FileDetails $Path   
  $EDGEVersion = $Meta.Comments.split(' ')[0]; 
  $Filename = (get-item $Path).name
  $destinationfolder = "$SCCM\$EdgeVersion";
  hl "Downloaded version: $EdgeVersion" $EdgeVersion 
  hl "Destination folder: $destinationfolder" "$destinationfolder" 
  SCCM-MapDrive
  IF (!(test-path $destinationfolder)) {
    hl "Creating $destinationfolder" "$destinationfolder"
    [System.IO.Directory]::CreateDirectory($destinationfolder); Write-Output "Moving $Path to $destinationfolder"  
    [System.IO.File]::Move($Path, "$destinationfolder\$Filename")  
  }
  $SavedPath = $(pwd)
  SCCM-LoadModule
  IF ((Get-CMDeploymentType -ApplicationName "Microsoft Edge $EdgeVersion" -DeploymentTypeName "DT_Edge_$EdgeVersion")) { hl "Already exist Microsoft Edge - $EdgeVersion" "Microsoft Edge - $EdgeVersion"; break }
  $NewApp = @{
    AppName              = "Microsoft Edge $EdgeVersion"
    Description          = "Microsoft Edge Installer"
    Publisher            = "Microsoft"
    SoftwareVersion      = $EdgeVersion
    Icon                 = "\\drscmsrv2\e$\SoftwarePackages\_ico\edge_browser.png"
    ContentLocation      = "$destinationfolder\$filename"
    InstallCommand       = "msiexec /i $filename /qn"
    DTName               = "DT_Edge_$EdgeVersion"
    FolderPath           = "DUB:\Application\Microsoft Edge"
    DPGroupName          = "AllDP"
    EstimatedRuntimeMins = 10 
  }
  SCCM-NewApp @NewApp
  <#
  SCCM-NewApp `
    -AppName     "Microsoft Edge $EdgeVersion" `
    -Description "Microsoft Edge Installer" `
    -Publisher   "Microsoft" `
    -SoftwareVersion $EdgeVersion `
    -Icon "\\drscmsrv2\e$\SoftwarePackages\_ico\edge_browser.png" `
    -ContentLocation "$destinationfolder\$filename" `
    -InstallCommand "msiexec /i $filename /qn" `
    -DTName "DT_Edge_$EdgeVersion" `
    -FolderPath "DUB:\Application\Microsoft Edge" `
    -DPGroupName "AllDP" `
    -EstimatedRuntimeMins 10
#>
  $grp = "Test_MB", "SCCM Pre-Test Group", "SCCM Test Group"
  $grp[0..1] | % { DeployToGroup -Apps "Microsoft Edge $EdgeVersion" -Collection $_ -Now -WhatIf }
  DeployToGroup -Apps "Microsoft Edge $EdgeVersion" -Collection 'SCCM Test Group' 
  $grp | % { Invoke-CMClientNotification -ActionType ClientNotificationRequestMachinePolicyNow -CollectionName $_ }
  cd $SavedPath 
} 

function Pack-Calypso([switch]$hex) {
  $fpath = '\\drscmsrv2\e$\SoftwarePackages\Calypso\'
  if ($hex) { $ex = '_hex' } else { $ex = '' }
  $fname = (gci "$fpath\*TR??$ex" | sort CreationTime -Descending | select -first 1).name  
  $TRver = ($fname | Select-String "TR(\d{2})").Matches.value 
  $path = Join-Path $fpath $fname   #$Ver = Split-Path $path -Leaf
  $AppName = "Calypso $fname"       #$Ver -match "\d*TR\d\d.*"; $tr = $Matches[0]; 
  cd c:
  $DetectionFile = gci "$path\client\TR*.txt" -Name
  "$path `nDetectionFile is $DetectionFile - Preparing package - $appname - $fname - $TRver"
  pause
  Import-Module (Join-Path $(Split-Path $env:SMS_ADMIN_UI_PATH) ConfigurationManager.psd1); cd "DUB:\"
  $newApp = @{ Name  = $AppName 
    Description      = "Calypso $fname $(get-date)" 
    Publisher        = 'AIB DRS'
    SoftwareVersion  = $fname
    IconLocationFile = "\\drscmsrv2\e$\SoftwarePackages\Calypso\calypso.png" 
  }
  $app = Get-CMApplication -Fast -Name $Appname
  if (!($app)) { $app = New-CMApplication @newApp }
  $app | Select LocalizedDisplayName, LocalizedDescription
  $addDT = @{ ApplicationName = $Appname
    DeploymentTypeName        = "DT_$Appname"
    InstallCommand            = 'powershell.exe -ExecutionPolicy Bypass -Command .\SCCM-CalypsoJob.ps1'
    ContentLocation           = $path
    InstallationBehaviorType  = 'InstallForSystem'
    EstimatedRuntimeMins      = 10
    LogonRequirementType      = 'WhetherOrNotUserLoggedOn'
    ScriptLanguage            = 'PowerShell'
    ScriptText                = ''
    Comment                   = "$(get-date) - $AppName"
  }
  Add-CMScriptDeploymentType @addDT | Select LocalizedDisplayName, LocalizedDescription
  $cl1 = New-CMDetectionClauseFile -FileName $DetectionFile -Path "%ProgramFiles%\CalypsoThickClient\client" -Existence
  Set-CMScriptDeploymentType -ApplicationName $Appname -DeploymentTypeName "DT_$Appname" -AddDetectionClause $cl1
  Move-CMObject -FolderPath "DUB:\Application\Calypso" -InputObject $app 
  Start-CMContentDistribution -InputObject $app -DistributionPointName 'drscmsrv2.dealers.aib.pri' -ErrorAction SilentlyContinue -DistributionPointGroupName 'AllDP'

  <#
unpack to cmsrv
run prep calypso on folders
add HomeUserFolder
add current prop file
add TRXX.txt for detection method
create a package in sccm
distribute
deploy to test calypso pc
same for TRxx_hex and add jstack2.bat to bin and start jstack2.bat to Navigator(Pre)Prod.bat
#>
}

function DeployToGroup {   
  param ( [string]$Apps = "$(Get-Date -f yyyy-MM)-*",
    [Parameter(Mandatory = $False)][ValidateSet('SCCM Pre-Test Group', 'SCCM Test Group', 'SCCM Group 1', 'SCCM Group 2', 'SCCM Group 3', 'SCCM Group 4', 'Test_MB')]
    [string]$Collection = 'SCCM Pre-Test Group',
    [datetime]$StartTime = (Get-Date -Hour 22 -Minute 03),
    [switch]$WhatIf, [switch]$Now   ) 
        
  if ($now) { $StartTime = Get-Date }
  Import-Module (Join-Path $(Split-Path $env:SMS_ADMIN_UI_PATH) ConfigurationManager.psd1); cd "DUB:\"
  $appsToDeploy = Get-CMApplication -Name $Apps | select -ExpandProperty LocalizedDisplayName 
  if (-not $AppsToDeploy) { Write-Host "No applications found matching the pattern '$Apps'" -ForegroundColor Yellow; return } else { $appsToDeploy }
  hl "Deploying applications to collection: $Collection" "$Collection" -nonewline 1; hl ", start time: $StartTime , increment 15 min" "$StartTime"
  if ($WhatIf) { break }

  foreach ($app in $appsToDeploy) {
    $NewDep = @{  ApplicationName = $app
      CollectionName              = $Collection
      AvailableDateTime           = $StartTime
      #DeadlineDateTime = $Time
      DeployAction                = "Install"                
      DeployPurpose               = "Required"
      UserNotification            = "DisplaySoftwareCenterOnly"
      SendWakeupPacket            = $true
      AllowRepairApp              = $true  
      PersistOnWriteFilterDevice  = $false
    } 
    $NewDep | ft
    $Result = New-CMApplicationDeployment @NewDep | Select-Object ApplicationName, CollectionName, StartTime
    Write-Host "Deployed: $($Result.ApplicationName) to $($Result.CollectionName) at $($Result.StartTime)" -ForegroundColor Cyan
    $StartTime = $StartTime.AddMinutes(15)
  }; c:
}

function Get-UDVariable {
  get-variable | where-object { (@(
        "FormatEnumerationLimit",
        "MaximumAliasCount",
        "MaximumDriveCount",
        "MaximumErrorCount",
        "MaximumFunctionCount",
        "MaximumVariableCount",
        "PGHome",
        "PGSE",
        "PGUICulture",
        "PGVersionTable",
        "PROFILE",
        "PSSessionOption"
      ) -notcontains $_.name) -and `
    (([psobject].Assembly.GetType('System.Management.Automation.SpecialVariables').GetFields('NonPublic,Static') | Where-Object FieldType -eq ([string]) | ForEach-Object GetValue $null)) -notcontains $_.name
  }
}

function Shared-pcs {
  $cts = Get-ADComputer -Filter { Description -like "CTS shared*" } -prop Description | select name, description
  $log = $cts.name | % { Logged-User $_ } 
  $log | ft
}

function cts($id) {
  $log = Get-CTSlogged
  $log | ? { $_.usr -like "*$id*" } 
}

function Get-CTSlogged {
  $cts = Get-ADComputer -Filter { Description -like "CTS shared*" } -prop Description | select name, description
  foreach ($pc in $cts) { 
    [pscustomobject]@{
      pc   = $pc.name
      desc = $pc.description
      usr  = (Get-UserProfile $pc.name | ? { $_ -notlike "dsk_*" }) -join ',' 
    } 
  }
}

function Get-ADou($name) {
  Get-ADOrganizationalUnit -Filter "Name -like '*$name*'" | ft Name, DistinguishedName 
}

function New-DealerUser($id) {
  $secpas = ConvertTo-SecureString -String "Fresh123!" -AsPlainText -Force
  $id | % { 
    $u = Get-ADUser $_ -Server prd.aib.pri -Properties GivenName, Surname, DisplayName, Initials, Description, mail
    New-ADUser -Path 'OU=Non Treasury Users,OU=DRS Win 10 Users,DC=dealers,DC=aib,DC=pri' -Enabled $true `
      -Name $u.Name -GivenName $u.GivenName -Surname $u.Surname -DisplayName $u.DisplayName `
      -Initials $u.Initials -Description $u.Description -EmailAddress $u.mail -AccountPassword $secpas 
  }
  $id | % { Get-ADUser $_ }
}

function Get-approved {
  $SusServer = 'DrsOpsMgr3'

  [void][reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration")
  $Wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer($SusServer, $false, 8530)

  $updateScope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
  $updatescope.ApprovedStates = [Microsoft.UpdateServices.Administration.ApprovedStates]::Any
  $updatescope.IncludedInstallationStates = [Microsoft.UpdateServices.Administration.UpdateInstallationStates]::All
  #$updates = $wsus.GetUpdates($updateScope)

  $updatescope.ApprovedStates = [Microsoft.UpdateServices.Administration.ApprovedStates]::LatestRevisionApproved
  $approvals = $wsus.GetUpdateApprovals($updatescope) | Select @{L = ’ComputerTargetGroup’; E = { $_.GetComputerTargetGroup().Name } },
  @{L = ’UpdateTitle’; E = { ($wsus.GetUpdate([guid]$_.UpdateId.UpdateId.Guid)).Title } }, GoLiveTime, AdministratorName, @{L = ’UpdateId’; E = { [guid]$_.UpdateId.UpdateId.Guid } } | ? { $_.ComputerTargetGroup -like "*Win 10*" }  #| sort-object -Property UpdateTitle -Unique | sort GoLiveTime | ft 

  [regex]::match($txt, 'KB(\d+)').value


  # $approvals | select UpdateTitle -Unique | sort UpdateTitle | measure
  # $approvals.count

}

function SetAppr {
  $SusServer = 'DrsCmSrv2'

  [void][reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration")
  $Srv = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer($SusServer, $false, 8530)
  $updateScope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
  $updatescope.ApprovedStates = [Microsoft.UpdateServices.Administration.ApprovedStates]::Any
  $updatescope.IncludedInstallationStates = [Microsoft.UpdateServices.Administration.UpdateInstallationStates]::All
  $Srv.GetUpdateCount($updateScope)
  $updates = $Srv.GetUpdates($updateScope)

  $approvals | % {
    $a = $_
    $u = $srv.GetUpdate($_.UpdateId)
    $grp = $srv.GetComputerTargetGroups() | ? { $_.name -eq $a.ComputerTargetGroup }
    "$($u.KnowledgebaseArticles) - $($grp.name)"
    $u.Approve('Install', $grp)
  }
 
  $srvappr = $srv.GetUpdateApprovals($updatescope) 
}
  
function RegChange {
  Set-RemoteReg -PC $env:COMPUTERNAME -HKEY CurrentUser -Path 'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -name ShowSecondsInSystemClock -value 1 -kind DWord
}

function Close-AllApps {
  Get-Process | ? { $_.MainWindowTitle -ne "" -and $_.Id -ne $PID -and $_.ProcessName -ne "explorer" } | Stop-Process -Force
}

function New-MyVM {
  $name = 'Win11-02'
  $mac = 'CC96E542BE9C'
  New-VM -Name $name -MemoryStartupBytes 16GB -NewVHDSizeBytes 50GB -Generation 1 -SwitchName 'Hyper-V Switch' -NewVHDPath "C:\ProgramData\Microsoft\Windows\Virtual Hard Disks\$name.vhdx" 
  Add-VMNetworkAdapter -VMName $name -IsLegacy $true -SwitchName 'Hyper-V Switch' -StaticMacAddress $mac
  Set-VMNetworkAdapter -VMName $name -StaticMacAddress $mac 
  Set-VMProcessor $name -Count 12
}

function Sync-File($f1, $f2, $fmask, $log = "c:\temp\logs\Sync.txt") {
  $opt = "/R:1 /W:1 /xo /NP /NS /NC /NFL" -split ' '
  robocopy.exe $f1 $f2 $fmask $opt >> $log
  robocopy.exe $f2 $f1 $fmask $opt >> $log
 (gc $log -Tail 690) | ? { $_.trim() -and $_ -notlike "   ROBOCOPY*" } | Out-File $log 
}

function Test-Tanium ($pcs, $restart = 0) {
  #$pcs = @('HKK0Y04-LON','dkk0y04-dub','6ns3mm2-dub')
  foreach ($pc in $pcs) {
    $ver = Get-InstalledApp $pc "*tanium client*" 
    $ser = Get-Service -ComputerName $pc "*win*" | select *
    [PSCustomObject]@{ PC = $pc; Ver = $ver.AppName; service = $ser.Status }
    if ($ser -and $ver -and $restart) { Get-Service -ComputerName $pc "*tanium client*" | Restart-Service -Verbose }
  }
}

function SendKeys($Name, $Keys) {
  # {ENTER} {TAB} ^Ctrl +Shift %Alt https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/sendkeys-statement
  $wsh = New-Object -Com wscript.shell;
  $ids = (ListWindows | ? { $_.MainWindowTitle -like "*$name*" }).id
  $wsh.AppActivate("$name");
  Sleep -m 300;
  $wsh.SendKeys("$keys")
  sleep -m 100
}

function Many-MSTSC($pc) {
  $apc = (Import-Csv "G:\AIB Software\_test\VM_vulns_abeup5mb_20250110.csv").pc | % { ($_ -split '\.')[0] } | ? { $_ -ne 'G4Z71K3-BCS' }
  $pc = $apc[3]

  $wsh = New-Object -Com wscript.shell;
  $pc | % { New-MSTSC $_; Sleep -m 500; $wsh.AppActivate('Windows Security'); Sleep -m 500; $wsh.SendKeys("$pas{ENTER}") }
  sleep 10
  $ids = (ListWindows | ? { $_.MainWindowTitle -like "*Remote Desktop*" }).id
  $ids | % { $wsh.AppActivate($_); sleep 1; $wsh.SendKeys("{ENTER}"); sleep 1 }
  sleep 5
  $pc | % { Update-Photoapp $_ }
}

function Update-Photoapp($pc) {
  # $pc = 'CDQ44K3-BCS'
  $pc
  Run-r $pc { $p = "C:\Temp\UpdatePhotoApp"; cd $p; $c = { Add-AppxProvisionedPackage -PackagePath $_.FullName -online -SkipLicense }; gci $path *.appx | % $c; gci $path *.msix* | % $c } 
  Run-r $pc { Get-AppxProvisionedPackage -Online | ? { $_.displayname -like "*photo*" } } | tee -Variable runr
  (($runr -split "`n")[1] -split ':')[1].trim()
  #if ($ver -like "2024*") { $id = $((Logged-User $pc).id); MLogoff $pc $id }
}

function MLogoff($pc, $ses) {
  logoff $ses /SERVER:$pc
}

function Test-t {
  $pc = 'HFQ44K3-BCS'  #Run-r $pc "Get-appxprovisionedpackage –online | ? {`$_.DisplayName }"
  $p = "C:\Temp\UpdatePhotoApp"; cd $p; $c = { Add-AppxProvisionedPackage -PackagePath $_.FullName -online -SkipLicense }; gci $path *.appx  | % $c; gci $path *.msix* | % $c
}

function Run-r {
  [CmdletBinding()] param( 
    [Parameter(Mandatory = $True)] [string]$Pc = '5MK0Y04-DUB', $cmd, [switch]$noesc
  )
  if (!$pc) { "No PC name provided"; break }; if (!(Aping $pc)) { 'Offline'; break }
  if (!$noprase) { $cmd = $cmd -replace '"', '\$&' } # $cmd = "Get-appxprovisionedpackage –online -Verbose | where-object {`$_.displayname -like \`"*Teams*\`" }"
  $rr = Run-Remote $pc "powershell -command `"Start-Transcript c:\Temp\logs\logsps.txt; ''; $cmd; Stop-Transcript;`""
  while (Get-Process -ComputerName $pc -id $rr.ProcessId -ErrorAction SilentlyContinue) { sleep -m 200 }
  $c = gc "\\$pc\c$\Temp\logs\logsps.txt" -Raw
 ((($c -split '[\r\n]+(?=Transcript started)')[1] -split '\*\*\*\*+')[0] -split "`n`r" | select -Skip 1 | Out-String ).Trim()
  Write-Verbose "PC = $pc"
  Write-Verbose $cmd 
  Write-Verbose "powershell -command `"Start-Transcript c:\temp\logs\logsps.txt; ''; $cmd; Stop-Transcript`""
  <#
  Run-r $pc { Get-AppxProvisionedPackage -Online | ? { $_.displayname -like "*teams*" } | Remove-AppxProvisionedPackage -Online }
  Run-r $pc { DISM /Online /Add-ProvisionedAppxPackage /PackagePath:"c:\Temp\MSTeams-x64-n.msix" /SkipLicense /LogPath:"c:\Temp\Logs\Dism.txt" }
  Run-r $pc { msiexec /i "c:\Temp\WIN.msi" ALLUSERS=1 /log c:\temp\logs\msiexec.txt }
  Run-r 7TZXGL2-DUB { $updateSession = new-object -com "Microsoft.Update.Session"; $updates=$updateSession.CreateupdateSearcher().Search($criteria).Updates;wuauclt /reportnow }
 #>
}

function Run-rc {
  [CmdletBinding()] param( 
    [Parameter(Mandatory = $True)] [string]$Pc = '5MK0Y04-DUB', $cmd, [switch]$noesc, $cred
  )
  if (!$pc) { "No PC name provided"; break }; if (!(Aping $pc)) { 'Offline'; break }
  if (!$noprase) { $cmd = $cmd -replace '"', '\$&' } # $cmd = "Get-appxprovisionedpackage –online -Verbose | where-object {`$_.displayname -like \`"*Teams*\`" }"
  $rr = Run-RemoteCred $pc "powershell -command `"Start-Transcript c:\Temp\logs\logsps.txt; ''; $cmd; Stop-Transcript;`"" -cred $cred
  while (Get-Process -ComputerName $pc -id $rr.ProcessId -ErrorAction SilentlyContinue) { sleep -m 200 }
  $c = gc "\\$pc\c$\Temp\logs\logsps.txt" -Raw
 ((($c -split '[\r\n]+(?=Transcript started)')[1] -split '\*\*\*\*+')[0] -split "`n`r" | select -Skip 1 | Out-String ).Trim()
  Write-Verbose "PC = $pc"
  Write-Verbose $cmd 
  Write-Verbose "powershell -command `"Start-Transcript c:\temp\logs\logsps.txt; ''; $cmd; Stop-Transcript`""
  <#
  Run-r $pc { Get-AppxProvisionedPackage -Online | ? { $_.displayname -like "*teams*" } | Remove-AppxProvisionedPackage -Online }
  Run-r $pc { DISM /Online /Add-ProvisionedAppxPackage /PackagePath:"c:\Temp\MSTeams-x64-n.msix" /SkipLicense /LogPath:"c:\Temp\Logs\Dism.txt" }
  Run-r $pc { msiexec /i "c:\Temp\WIN.msi" ALLUSERS=1 /log c:\temp\logs\msiexec.txt }
  Run-r 7TZXGL2-DUB { $updateSession = new-object -com "Microsoft.Update.Session"; $updates=$updateSession.CreateupdateSearcher().Search($criteria).Updates;wuauclt /reportnow }
 #>
}

function PraseNetUse($netuse = (net use)) {
  $netuse -like '* \\*' | % { $Status, $Local, $Remote, $Null = $_ -split ' +', 4
    [PSCustomObject]@{
      Status = $Status
      Local  = $Local
      Remote = $Remote 
    } }
}

function start-TS($pc = 'drs2019test1') {
  $rpath = "\\$pc\C$\_ScanUpdates\"
  $lpath = "C:\_ScanUpdates\"
  if (-not (Test-Path $rpath)) {
    $null = mkdir $rpath -Force
    copy $lpath* $rpath -Include "*.cab", '*.ps1' -Force -Recurse 
  }
  $opt = New-CimSessionOption -Protocol DCOM
  $s = New-CimSession -Computername $pc -SessionOption $opt -OperationTimeoutSec 3 -ErrorAction Stop #-Credential $cred
  $action = New-ScheduledTaskAction -Execute powershell -Argument "-executionpolicy bypass -file $lpath\ScanUpdates.ps1" -WorkingDirectory $lpath 
  $principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -RunLevel Highest -LogonType S4U 
  if (Get-ScheduledTask 'Test-Scan' -CimSession $s) { Unregister-ScheduledTask 'Test-Scan' } 
  #$newTS = New-ScheduledTask -Action $action -Description 'Test remote Task' -Principal $principal  #Register-ScheduledTask 'Test-Scan' -InputObject $newTS -CimSession $s
  Register-ScheduledTask 'Test-Scan' -Action $action -Description 'Test remote Task' -Principal $principal -CimSession $s
  Get-ScheduledTask -CimSession $s -TaskName 'Test-Scan' | Start-ScheduledTask 
  Remove-CimSession $s 
}

function Get-PcInfoDesktops {
  $out = $adc | % { $pc = $_.Name
    [PSCustomObject]@{ Hostname = $_.name;
      Model                     = Get-Model $pc
      Serial                    = (Get-WmiObject Win32_bios -ComputerName $pc).SerialNumber
      Ip                        = (aping $pc).address.IPAddressToString
      Mac                       = (Get-Mac $pc).MACAddress -join ', '
      Ver                       = (Get-Winver $pc).ver
    }
  } 
  Export-Xlsx -Path C:\Users\dsk_58691\Desktop\mm.xlsx -obj $out
}

function Get-Drama {
  $year = Get-Date –f yyyy
  $qtr = [math]::Ceiling((Get-Date).Month / 3) 
  $file = "G:\DRAMA\$year\$($year)Q$($qtr)\DRAMA Dashboard $year Q$qtr.xlsx"
  Import-Excel $file -WorksheetName 'email log' | tee -Variable global:Drama
}

function secStr($s) {

  $m = ConvertTo-SecureString 'Dr11ms11' -AsPlainText -Force
  $kod = $m | ConvertFrom-SecureString
  $q = $kod | ConvertTo-SecureString
  $q.Length

  $SecurePassword = ConvertTo-SecureString 'Dr11ms11' -AsPlainText -Force
  $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
  $UnsecurePassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
  [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)

}

function esc($text) {
  [Management.Automation.WildcardPattern]::Escape($text)
  [regex]::Escape($text)
} 

function Init {
  #$ErrorActionPreference='silentlycontinue'
  $global:ModulePath = 'H:\MB\PS\modules\MBMod\0.3\' 
  $global:ModulePath1 = $PSCommandPath
  $global:ModulePath2 = (Get-Module -Name mbmod).ModuleBase
  $global:ModuleDir = if ($ModulePath2) { Split-Path (Split-Path $ModulePath2) }
  $global:ScriptFile = Get-CallingFileName
  $global:ScriptPath = if ($x = Get-CallingFileName) { Split-Path $x } else { ScriptDir }
  $global:upath = "$ModuleDir\MBMod\0.3\users.xlsx"
  $global:cpath = "$ModuleDir\MBMod\0.3\comps.xlsx"
  $global:DesktopPath = [Environment]::GetFolderPath("Desktop")
  $global:PatternSID = 'S-1-5-21-\d+-\d+\-\d+\-\d+$'

  #"Mß v1.3"
}

function Main {
  $global:ScriptPath = if ($psise) { Split-Path $psise.CurrentFile.FullPath } else { $PSScriptRoot }
}

Main

function Change-Password {
  #explorer.exe shell:::{2559a1f2-21d7-11d4-bdaf-00c04f60b9f0}
(New-Object -ComObject "Shell.Application").WindowsSecurity()
}

function Test-Modules {
  Init
  $path = "$ModuleDir"  #"$ScriptPath\modules"
  #if ($ScriptPath -eq $ModulePath2) {$path = Split-Path (Split-Path $ScriptPath)}
  $modUNC = @{ ImportExcel = "$path\ImportExcel\7.4.1\ImportExcel.psd1"
    MSCatalog              = "$path\MSCatalog\MSCatalog.psd1"
  }
  $ModUNC.keys | % { If (-not(Get-module $_)) { Import-Module $($ModUNC[$_]) -Global -WA SilentlyContinue } }
}

function Me-Import {
  Import-Module "H:\MB\PS\modules\MBMod\0.3\MBMod.psm1" -WA SilentlyContinue -Force -Global
  #Import-Module "$ScriptPath\modules\MBMod\0.3\MBMod.psm1" -Force -Global -WarningAction SilentlyContinue
  Init
}

function ImportMe {
  #iex ${using:function:ImportMe}.Ast.Extent.Text;ImportMe
  Import-Module "H:\MB\PS\modules\MBMod\0.3\MBMod.psm1" -WA SilentlyContinue -Force
  #Import-Module "$ScriptPath\modules\MBMod\0.3\MBMod.psm1" -Force -Global -WarningAction SilentlyContinue
  Init
}


function Get-CallingFileName {
  $cStack = @(Get-PSCallStack | ? { $_.ScriptName -and $_.ScriptName -notlike "*MBMod.psm1*" } )
  $cStack.ScriptName
}

function ScriptDir {
  #Only in local file
  if ($psise) { Split-Path $psise.CurrentFile.FullPath } else { $PSScriptRoot }
  #$global:GetScriptDir = { if ($psise) {Split-Path $psise.CurrentFile.FullPath} else {$PSScriptRoot} }
} 

function ListWindows {
  Get-Process | Where { $_.MainWindowTitle } | Select-Object ProcessName, MainWindowTitle
}

function Test-BCS {
  $OutFile = "Central Park Checks $(get-date -Format 'yyyy-MM-dd HH-mm').xlsx" 
  $OutPath = 'G:\Daily Checks\Completed Central Park Checks'

  Test-Modules;
  $path = Join-Path $OutPath $OutFile
  "Output to excel file $path"

  $inCP = Get-ADComputer -Filter * -SearchBase 'OU=DRS Central Park,OU=DRS Win 10 PCs,DC=dealers,DC=aib,DC=pri' -Properties description, location
  $all = Get-ADComputer -Filter * -Properties description, location
  $list = $inCP.name  #(Get-ADComputer -Filter { OperatingSystem -NotLike "*server*" } ).name

  rv ii -ErrorAction SilentlyContinue 

  $out = foreach ($pc in $list) {
    MyProgress $pc $list.count
    [PSCustomObject]@{ 
      PC          = $pc;
      AD          = $ad = $pc -in $all.name   
      #SCCM = $sccm = if ((Get-CMCollectionOfDevice $pc) -like "*does not exist in Site*") {$false} else {$true}
      Ping        = $ping = [bool](APing $pc)
      WMI         = $wmi = if ($ping) { [bool](Check-WMI $pc -timeout 5) } else { $false }
      Pass        = $AD -and $Ping -and $wmi 
      Location    = ($all | ? { $_.name -eq $pc }).Location
      Description = ($all | ? { $_.name -eq $pc }).Description
    }
  }

  $out | ft
  "$(($out.pass -eq $true).count) out of $($list.count) computers online in Central Park"
  $path
  $out | ? { ! $_.pass } | ft
  $global:BCS = $out

  #$c1 = New-ExcelChartDefinition -YRange "PC" -XRange "Pass" -Title "Total"  -NoLegend -Height 225 -Row 9  -Column 15
  $o = Export-Excel -PassThru -NoNumberConversion Name -Path $path -InputObject ($out | select PC, Pass) -TableName 'Summary' -WorksheetName 'Summary' -FreezeTopRow -BoldTopRow -AutoSize -CellStyleSB { param($workSheet)  $WorkSheet.Cells.Style.HorizontalAlignment = "Left" } #`-Barchart -ExcelChartDefinition $c1
  Add-ConditionalFormatting -Worksheet $o.Summary -Range "B2:B52" -RuleType ContainsText -ConditionValue "TRUE" -ForegroundColor Green -BackgroundColor LightGreen
  Add-ConditionalFormatting -Worksheet $o.Summary -Range "B2:B52" -RuleType ContainsText -ConditionValue "FALSE" -ForegroundColor Red -BackgroundColor LightPink
  $o.Summary.Cells["D1"].Value = "$(($out.pass -eq $true).count)"
  $o.Summary.Cells["E1"].Value = "out of"
  $o.Summary.Cells["F1"].Value = "$($list.count)"
  if (($out.pass -eq $true).count -eq $list.count) { $color = [System.Drawing.Color]::LightGreen } else { $color = [System.Drawing.Color]::LightPink } 
  Set-ExcelRange -Worksheet $o.Summary -Range "D1:F1" -BackgroundColor $color 
  $o.Summary.Cells["H1"].Value = 'Date'
  $o.Summary.Cells["H2"].Value = 'Tested By'
  $o.Summary.Cells["H3"].Value = 'Manager Signed'
  $o.Summary.Cells["I1"].Value = (Get-Date -format g)
  $x = (Get-aduser ($env:USERNAME -replace 'dsk_' -replace 'adm_'))
  $o.Summary.Cells["I2"].Value = "$($x.GivenName) $($x.Surname)"
  Set-ExcelRange -Worksheet $o.Summary -Range "H1:H3" -BackgroundColor SkyBlue -BorderAround Thin 
  Set-ExcelRange -Worksheet $o.Summary -Range "H1:I3" -BorderAround Thin
  Set-ExcelRange -Worksheet $o.Summary -Range "H1:I2" -BorderAround Thin -AutoSize
  $o.Summary.Cells.AutoFitColumns()
  $o.Summary.Cells["D1:F1"].AutoFitColumns(3)
  $o = Export-Excel -ExcelPackage $o -WorksheetName "Details" -InputObject $out -Show -TableName 'Details' -FreezeTopRow -BoldTopRow -AutoSize -PassThru
  Add-ConditionalFormatting -Worksheet $o.Details -Range "B2:E52" -RuleType ContainsText -ConditionValue "TRUE" -ForegroundColor Green -BackgroundColor LightGreen
  Add-ConditionalFormatting -Worksheet $o.Details -Range "B2:E52" -RuleType ContainsText -ConditionValue "FALSE" -ForegroundColor Red -BackgroundColor LightPink
  if ($show) { Export-Excel -ExcelPackage $o -Worksheet $o.Details -Show -AutoSize }
  else { Export-Excel -ExcelPackage $o -Worksheet $o.Details -AutoSize }

  # ii $OutPath
}

function Get-UTime { [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds() }

function SetCursor($x, $y) { $Host.UI.RawUI.CursorPosition = @{x = $x; y = $y } }

function GetCursor { $Host.UI.RawUI.CursorPosition -split ',' }

function Wysw($x, $y, $text, $ftime = 75, $color) {
  $ox, $oy = GetCursor
  SetCursor $x $y
  if ($color) { Write-Host $text -NoNewline -ForegroundColor $color } 
  else { Write-Host $text -NoNewline }
  SetCursor $ox $oy
}

function Zegar {
  Wysw 60 1 $(Get-Date -F 'HH:mm:ss') 
}

function Spinner($x, $y, $speed = 75) {
  $anim = "|/-\".ToCharArray(); 
  $i = [Math]::Round((Get-UTime) / $speed, 0) % $anim.count
  $ox, $oy = $Host.UI.RawUI.CursorPosition -split ','
  $Host.UI.RawUI.CursorPosition = @{x = $x; y = $y }
  write-host "$($anim[$i])`b" -NoNewline -ForegroundColor Green
  $Host.UI.RawUI.CursorPosition = @{x = $ox; y = $oy }
}

function Test-ADCMWsusEpo {
  if ($ExecutionContext.SessionState.LanguageMode -ne 'FullLanguage') { 'Please start script with administrator right'; pause; exit }

  $w = Get-WsusServer -Name drsopsmgr3 -PortNumber 8530
  $WSUS = Get-WsusComputer -UpdateServer $w | ? { $_.ComputerRole -eq 'Workstation' }
  $WSUSList = ($WSUS.FullDomainName -replace '.dealers.aib.pri').ToUpper()

  $w2 = Get-WsusServer -Name drscmsrv2 -PortNumber 8530
  $WSUS2 = Get-WsusComputer -UpdateServer $w2 | ? { $_.ComputerRole -eq 'Workstation' }
  $WSUSList2 = ($WSUS2.FullDomainName -replace '.dealers.aib.pri').ToUpper()

  Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" -ErrorAction SilentlyContinue -Force; cd "DUB:\"
  $CM = Get-CMDevice | ? { $_.IsClient } | select name, LastActiveTime, MACAddress
  $CMList = $CM.name.ToUpper(); C:
  $AD = (Get-ADComputer -Filter { OperatingSystem -NotLike "*server*" } -Properties Description, CanonicalName, MemberOf, Location) | ? { $_.name -ne 'DRSVCENTRE' } | select *
  $ADList = $AD.Name.ToUpper()
  $last_ePO_file = Get-Item "G:\documentation and procedures\Vulnerability Management\EPO 30 Day Reviews\EPO_*.csv" | sort CreationTime -Descending | select -first 1 

  $Epo = Import-Csv $last_ePO_file | ? { $_.'OS Platform' -eq 'Workstation' } | 
  % { $lc = $_.'Last Communication' -replace ' BST' -replace ' GMT'; 
    [PSCustomobject]@{ System = 'ePO'; Name = $_.'System Name'; LastComm = if ($lc -match "AM|PM") { [datetime]$lc } else { [datetime]::ParseExact($lc, "dd/MM/yy HH:mm:ss", $null) }
      Data = $_.'IP address'; User = $_.'User Name'; Managed = $_.'Managed State'; VerAgent = $_.'Product Version (Agent)'; VerEndpoint = $_.'Product Version (Endpoint Security Platform)'; OS = $_.'OS Platform'
    } }
  $EpoList = $Epo.Name  
 
  $Total = $ADList + $CMList + $WSUSList + $EpoList | select -Unique
  $all = $global:DRS = $Total | % { 
    $pc = $_; 
    $iad = ($AD | ? { $_.Name -like "*$pc*" }); 
    $iepo = ($Epo | ? { $_.Name -like "*$pc*" }); 
    $icm = ($CM | ? { $_.Name -like "*$pc*" }); 
    $iwsus = ($wsus | ? { $_.FullDomainName -like "*$pc*" })
    [PSCustomObject]@{ PC = $pc; Desc = $iad.Description 
      WSUS_LastComm = $wsus_lc = $iwsus.LastReportedStatusTime;
      EPO_LastComm = $epo_lc = $iepo.LastComm;
      CM_LastComm = $icm.LastActiveTime
      IP = if ($wsus_lc -gt $epo_lc) { $iwsus.IPAddress } else { $iepo.Data }
      MAC = $icm.MACAddress
      MemberOf = $iad.MemberOf -replace "CN=|,DC=dealers,DC=aib,DC=pri|,OU=SCCM Computer Groups|,OU=Dublin|,OU=BCM" -join ', '
      Cannon = $iad.CanonicalName -replace "dealers.aib.pri|$pc"
      AD = ($_ -in $ADList); CM = ($_ -in $CMList); WSUS = ($_ -in $WSUSList); ePO = ($_ -in $EpoList); 
    } } 

  [array]$comps = $global:DRSreport = $all | ? { !$_.Ad -or ! $_.CM -or ! $_.ePO -or ! $_.WSUS -or
    $_.WSUS_LastComm -lt (Get-Date).AddDays(-10) -or
    $_.EPO_LastComm -lt (Get-Date).AddDays(-10) -or
    $_.CM_LastComm -lt (Get-Date).AddDays(-10) } | select * 
  $comps = ($comps | sort pc | ft | Out-String).Trim()

  if (-not $comps) { $comps = 0 }                    
  $out = @"
Number of PCs in AD   : $($ADList.Count)
Number of PCs in CM   : $($CMList.Count)
Number of PCs in ePO  : $($EpoList.Count)
Number of PCs in WSUS : $($WSUSList.Count)

Computers that have not called in to the systems for more than 10 days or are missing from any of the systems
$($comps | Out-String)

"@
  $out 

  if ($nosave) { break }
  $fname = "G:\documentation and procedures\Vulnerability Management\EPO 30 Day Reviews\Dealers_report_$(Get-Date -Format 'yyyy-MM-dd_HH-mm')"
  $out | Out-File "$fname.txt"
  #$all | Export-Csv "$fname.csv" -NoTypeInformation
  Test-Modules
  Export-xlsx $all "$fname.xlsx" 
  "$fname.txt"

  # ii "G:\documentation and procedures\Vulnerability Management\EPO 30 Day Reviews\"
}

function Close-File($pc = 'DrsCorpSrv2', $name = "*xls*", $user = "*", $ReallyClose = 0) {
  #*Activity Report
  if (Test-Path variable:cred) { if ($cred.UserName -ne 'adm_58691') { $cred = Get-Credential adm_58691 } } else { $cred = Get-Credential adm_58691 }
  $opt = New-CimSessionOption -Protocol DCOM
  $s = New-CimSession -Computername $pc -SessionOption $opt -ErrorAction Stop -Credential $cred
  Get-SmbOpenFile -CimSession $s -ClientUserName $user | ? { $_.Path -like $name }  | tee -Variable OpenFiles | select ClientUserName, Path
  if ($ReallyClose) { "The files will be closed"; pause; $OpenFiles | Close-SmbOpenFile }
  Remove-CimSession $s 
}

function SIDtoUser($SID) {
  $SID.Translate([System.Security.Principal.NTAccount])
}

function Get-Licence($pc, $app = 'Windows%') {
  $opt = New-CimSessionOption -Protocol DCOM
  $s = New-CimSession -Computername $pc -SessionOption $opt -ErrorAction Stop
  Get-CimInstance SoftwareLicensingProduct -CimSession $s -Filter "Name like '$app'" | where { $_.PartialProductKey } | select *
  Remove-CimSession $s 
}

function Get-DNSsuffix($pc) {
  $opt = New-CimSessionOption -Protocol DCOM
  $s = New-CimSession -Computername $pc -SessionOption $opt -ErrorAction Stop
  Get-DnsClient -CimSession $s | ft
  Get-DnsClientServerAddress -CimSession $s | ft
  Remove-CimSession $s 
}

function WSUS-ForceUpdate {
  $updateSession = new-object -com "Microsoft.Update.Session"; $updates = $updateSession.CreateupdateSearcher().Search($criteria).Updates
  #Running this commands will "prime" the Windows Update engine to submit its most recent status on the next poll.  To trigger that next poll, use:
  wuauclt /reportnow
}

function EnableADAL($pc, $usr) {
  test-path "\\$pc\c$\users\$usr\ntuser.dat"
 (isLogged $pc).user
  reg load "HKU\$pc-$usr" "\\$pc\c$\users\$usr\ntuser.dat"
  Set-ItemProperty "Registry::HKEY_USERS\$pc-$usr\SOFTWARE\Microsoft\Office\16.0\Common\Identity" -Name EnableADAL -Value 1 -Verbose
  $null = REG UNLOAD "HKU\$pc-$usr"
}

function Get-PhysicalDiskR($PC) {
  $opt = New-CimSessionOption -Protocol DCOM
  try {
    $s = New-CimSession -Computername $pc -SessionOption $opt -ErrorAction SilentlyContinue 
    Get-PhysicalDisk -CimSession $s
    Remove-CimSession $s 
  }
  catch { $false } 
}

<#
ADinfo
$l = Ping-DealersPCs
$out = $l |% { $_; $temp=(Get-PhysicalDiskR $_); [PSCustomObject]@{ PC = $_; MediaType=$temp.MediaType; FriendlyName=$temp.FriendlyName; SerialNumber=$temp.SerialNumber; Size=$temp.Size  } }
Export-Desktop $out 'SSD'
#>

function KMS($pc) {
  #/ckms
  $all = cscript.exe "$env:SystemRoot\System32\slmgr.vbs" "$pc" /dli
  $LicStatus = (($all | ? { $_ -match 'Volume activation expiration:' }) -split ':')[1].Trim()
  $KMS = (($all | ? { $_ -match 'Registered KMS machine name:' }) -split ':')[1].Trim()
  $KMS_DNS = (($all | ? { $_ -match 'KMS machine name from DNS:' }) -split ':')[1].Trim()
  # return an object
  [PsCustomObject]@{
    ComputerName  = $pc
    LicenseStatus = $LicStatus
    KMS           = $KMS
    minutes       = ($LicStatus -split ' ')[0]
    KMS_DNS       = $KMS_DNS
    all           = $all -join "`n`t"
  }
}

function Get-Model($pc) {
 (Get-WmiObject Win32_ComputerSystem -ComputerName $pc).Model
}

function Get-EdgeDriver($dir = 'C:\Selenium\src') {
  #[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
  $key = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{56EB18F8-B008-4CBD-B6D2-8C97FE7E9062}'
  $version = (Get-ItemProperty -Path $key -Name pv).pv
  $ver = ($version -split '\.')[0]
  $type = 'edgedriver_win64'
  $url = "https://msedgedriver.azureedge.net/$version/$type.zip"
  Invoke-WebRequest -Uri $url -OutFile "$dir\$type.zip"
  Expand-Archive -LiteralPath "$dir\$type.zip" -DestinationPath "$dir\$ver"
  Remove-Item "$dir\$type.zip"
}

function Get-DesktopUpdates {
  # Save-MSCatalogUpdate
  Test-Modules
  Set-Proxy 1
  Sleep -Seconds 2
  $Last30Days = { $_.LastUpdated -gt (Get-Date).AddDays(-20) }
  $d = @(Get-MSCatalogUpdate -Search "Cumulative*Windows 10*22H2*x64" -Strict -ExcludePreview | ? { $_.Title -notlike "*Dynamic*" -and $_.Title -notlike "*4.8.1*" } | ? $Last30Days)
  #$d += Get-MSCatalogUpdate -Search "Update*2016*32" -Strict | ? $Last30Days
  #$d += Get-MSCatalogUpdate -Search "Update*2016*64" -Strict | ? $Last30Days
  $d | sort Title -Unique | tee -Variable global:kbs 
  Set-Proxy 0
}

function Get-ServerUpdates {
  Test-Modules
  Set-Proxy 1
  Sleep -Seconds 1
  $Last30Days = { $_.LastUpdated -gt (Get-Date).AddDays(-30) }
  $o = @( Get-MSCatalogUpdate -Search (Get-Date –f yyyy-MM) -AllPages | ? $Last30Days )
  $o += Get-MSCatalogUpdate -Search "Cumulative Update for Windows Server 2012 R2" -AllPages | ? $Last30Days 
  $o += Get-MSCatalogUpdate -Search "Cumulative Update for Windows Server 2016" -AllPages | ? $Last30Days 
  $o += Get-MSCatalogUpdate -Search "Cumulative Update for Windows Server 2019" -AllPages | ? $Last30Days
  $o += Get-MSCatalogUpdate -Search "Security Monthly Quality Rollup" -AllPages | ? $Last30Days
  #$o += Get-MSCatalogUpdate -Search "SQL" -AllPages  | ? $Last30Days
  $o += Get-MSCatalogUpdate -Search "Servicing Stack Update for Windows Server*x64" -Strict -SortBy Products -AllPages | ? $Last30Days
  $o | ? { $_.Products -in @("Windows Server 2012 R2", "Windows Server 2016", "Windows Server 2019") } | sort Title -Unique | sort Products | tee -Variable global:SrvKB
  Set-Proxy 0
}

function New-MSPapp($a) {
  $YearMonth = Get-Date –f yyyy-MM; $qtr = [math]::Ceiling((Get-Date).Month / 3) 
  $path = "\\drscmsrv2\e$\SoftwarePackages\Monthly Patches\$YearMonth"
  $CMFolder = "DUB:\Application\_Security Update"
  cd c:
  $files = gci $path\*.msp -Recurse 
  foreach ($file in $files) {
    $i = ($file.name | Select-String "^(KB\d+)-(\d{4})-(\d{2})-(\w+)-x-none").Matches.Groups.value 
    $Info = "$($i[1])-$($i[2])-$($i[4])"
    $Appname = "$(Get-date -f "yyyy-MM")-$Info"; "`n$Appname-$($i[3])"
    $bitness = if ($i[3] -eq 32) { 'x86' } else { 'x64' }
    Import-Module (Join-Path $(Split-Path $env:SMS_ADMIN_UI_PATH) ConfigurationManager.psd1); Set-Location -Path "DUB:\"
    $newApp = @{ Name = "$Appname"
      Description     = $file.name 
      Publisher       = 'Microsoft'
      SoftwareVersion = "$YearMonth" 
    }
    $app = Get-CMApplication -Name $Appname
    if (!($app)) { $app = New-CMApplication @newApp } else { 'App exists ' }
    $app | Select LocalizedDisplayName, LocalizedDescription
    $script = @'
$KBNumber = "KB_Number"
$bits = "Bit_Number"
$RegPath = @("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
If (${Env:ProgramFiles(x86)}){$RegPath += "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"}
$KBList = Get-ChildItem -Path $RegPath -Recurse | Get-ItemProperty | Where-Object {$_.DisplayName -like "*$KBNumber*$bits*"} 
If ($KBList){$KBList | % {Write-Host `n"$($_.DisplayName) found!"}}
'@ -replace 'KB_Number', $i[1] -replace "Bit_Number", $i[3]
    $addMsp = @{ ApplicationName = $Appname
      DeploymentTypeName         = "DT_$Appname-$($i[3])"
      InstallCommand             = "msiexec.exe /p $($file.Name) /qn"
      ContentLocation            = "$($file.Directory)"
      InstallationBehaviorType   = 'InstallForSystem'
      EstimatedRuntimeMins       = 10
      LogonRequirementType       = 'WhetherOrNotUserLoggedOn'
      ScriptLanguage             = 'PowerShell'
      ScriptText                 = $script
      Comment                    = "$(get-date) - $($file.Name)"
      AddRequirement             = Get-CMGlobalCondition -Name "*Office bitness*" | New-CMRequirementRuleCommonValue -Value1 $bitness -RuleOperator IsEquals
    }
    if (!(Get-CMDeploymentType -InputObject $app -DeploymentTypeName $addMsp.DeploymentTypeName)) {
      Add-CMScriptDeploymentType @addMsp | Select LocalizedDisplayName, LocalizedDescription
      if ($file.name -notin $app.LocalizedDescription) { Set-CMApplication -InputObject $app -Description "$($app.LocalizedDescription), $($file.name)" }  
    }
    else { 'Deployment exists' }
    if (-not (Test-Path "$CMFolder\$YearMonth")) { New-CMFolder -Name $YearMonth -ParentFolderPath $CMFolder }
    Move-CMObject -FolderPath "$CMFolder\$YearMonth" -InputObject $app 
    $app = Get-CMApplication -Name $Appname
    if ((Get-CMDistributionStatus -Id ($app.PackageID) -ErrorAction SilentlyContinue).Targeted -eq 0) {
      Start-CMContentDistribution -InputObject $app -DistributionPointName 'drscmsrv2.dealers.aib.pri' -ErrorAction SilentlyContinue 
    }
  }
}

function New-MSUapp {
  # net use Y: \\drscmsrv2\e$ /USER:adm_58691 *
  $YearMonth = Get-Date –f yyyy-MM; $qtr = [math]::Ceiling((Get-Date).Month / 3) 
  $path = "\\drscmsrv2\e$\SoftwarePackages\Monthly Patches\$YearMonth"
  $CMFolder = "DUB:\Application\_Security Update"

  cd c:
  $files = gci $path\*.msu -Recurse #| select -First 2  
  foreach ($file in $files) {
    $i = ($file.name | Select-String "^windows10.0-(KB\d+)-(x\d{2})").Matches.Groups.value ; $i
    $cunet = if ($file.name -match "ndp\d{2}") { 'NET' } else { 'CU' }
    $info = "$($i[1])-$($i[2])-$cunet"
    $Appname = "$(Get-date -f "yyyy-MM")-$Info"; "`n" + $Appname
    Import-Module (Join-Path $(Split-Path $env:SMS_ADMIN_UI_PATH) ConfigurationManager.psd1); cd "DUB:\"
    $newApp = @{ Name = "$Appname"
      Description     = $file.name
      Publisher       = 'Microsoft'
      SoftwareVersion = "$YearMonth $info" 
    }
    $app = Get-CMApplication -Fast -Name $Appname
    if (!($app)) { $app = New-CMApplication @newApp }
    $app | Select LocalizedDisplayName, LocalizedDescription
    $script = 'get-hotfix | Where-Object {$_.HotFixID -match "' + $i[1] + '"}'
    $addMsu = @{ ApplicationName = "$Appname"
      DeploymentTypeName         = "DT_$Appname"
      InstallCommand             = "$($file.Name) /quiet"
      ContentLocation            = "$($file.Directory)"
      InstallationBehaviorType   = 'InstallForSystem'
      EstimatedRuntimeMins       = 10
      LogonRequirementType       = 'WhetherOrNotUserLoggedOn'
      ScriptLanguage             = 'PowerShell'
      ScriptText                 = $script
      Comment                    = "$(get-date) - $($file.Name)"
    }
    if (!(Get-CMDeploymentType -InputObject $app -DeploymentTypeName $addMsu.DeploymentTypeName)) {
      Add-CMScriptDeploymentType @addMsu | Select LocalizedDescription, LocalizedDisplayName
      Set-CMApplication -InputObject $app -Description "$($app.LocalizedDescription), $($file.name)" 
    }
    if (-not (Test-Path "$CMFolder\$YearMonth")) { New-CMFolder -Name $YearMonth -ParentFolderPath $CMFolder }
    Move-CMObject -FolderPath "$CMFolder\$YearMonth" -InputObject $app 
    Start-CMContentDistribution -InputObject $app -DistributionPointName 'drscmsrv2.dealers.aib.pri' -ErrorAction SilentlyContinue 
  }
}

function Save-NewUpdate($path = "C:\Temp\updates") {
  if (-not (Test-Path $path)) { mkdir $path }
  Set-Proxy 1
  $u = (Get-DesktopUpdates)
  Set-Proxy 1
  $u | % { Save-MSCatalogUpdate $_ $path -AcceptMultiFileUpdates -UseBits -ErrorAction SilentlyContinue } 
  Set-Proxy 0
}

function ExtractCabsFolder ($CabFolder = 'C:\Temp\updates') {
  $files = gci "$CabFolder\*.cab"
  cd $CabFolder
  $UpFolder = (Split-Path $CabFolder)
  New-Item 'MSP' -ItemType Directory -force | Out-Null
  New-Item 'CabsDone' -ItemType Directory -force | Out-Null

  $msp = $CabFolder + '\MSP' 
  $CabsDone = $CabFolder + '\CabsDone' 
  foreach ($f in $files) {
    New-Item $f.BaseName -ItemType Directory -Force -Verbose | Out-Null
    $dir = $CabFolder + '\' + $f.BaseName
    expand $f.Name -F:*.msp $dir | Out-Null
    $a = gci "$dir\*.msp"
    if ($a.count -eq 1) { Rename-Item $a.Fullname "$($f.BaseName).msp" | Out-Null }
    Move-Item "$dir\*.msp" $msp -Force
    Move-Item "$f" $CabsDone -Force
    Remove-Item $dir -force 
  } 
}

function Move-toCM($path = 'C:\Temp\updates\') {
  $YearMonth = Get-Date –f yyyy-MM
  $qtr = [math]::Ceiling((Get-Date).Month / 3) 
  $pathCM = "\\drscmsrv2\e$\SoftwarePackages\Monthly Patches\$YearMonth"  #$msu = (Get-Item -Path $path\*.msu)

  foreach ($file in $msp = gci $path\*.msp -Recurse ) {
    # | select -skip 1 -First 1 $files = gci $path\*.msp -Recurse
    $i = ($file.name | Select-String "^(KB\d+)-(\d{4})-(\d{2})-(\w+)-x-none").Matches.Groups.value #$winv = ($file.name | Select-String -Pattern "\d{2}H2").Matches.Value.ToUpper()
    $Info = "$($i[1])-$($i[3])-$($i[2])-$($i[4])"
    $Appname = "$(Get-date -f "yyyy-MM")-$Info"; "`n" + $Appname
    md "$pathCM\$info" -Force
    move $file "$pathCM\$info"
  }

  foreach ($file in $msu = gci $path\*.msu -Recurse) {
    $i = ($file.name | Select-String "^windows10.0-(KB\d+)-(x\d{2})").Matches.Groups.value ; $i
    $cunet = if ($file.name -match "ndp\d{2}") { 'NET' } else { 'CU' }
    $info = "$($i[1])-$($i[2])-$cunet"
    $Appname = "$(Get-date -f "yyyy-MM")-$Info"; "`n" + $Appname
    md "$pathCM\$info" -Force
    move $file "$pathCM\$info"
  }
}

function SHowFaultyDeployment {
  Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1"; cd DUB:
  $appsToDeploy = Get-CMApplication -Fast -Name "2023-10-*"   # | select -ExpandProperty LocalizedDisplayName 
  $appsToDeploy | % { Get-CMApplicationDeploymentStatus -InputObject $_ } | Get-CMDeploymentStatusDetails  | % { [PSCustomObject]@{ PC = $_.MachineName; EnforcementState = $_.EnforcementState; AppName = $_.AppName; StatusType = $_.StatusType } } | ? { $_.EnforcementState -ne 1000 } 

}

function Get-NewUpdateSCCM {
  Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1"
  $SavedPath = $(pwd)
  cd DUB:
  Get-CMSoftwareUpdate -fast -DatePostedMin '07/09/2023' -IsLatest $true | 
  ? { $_.LocalizedDisplayName -like "*2016*32-Bit*" -or $_.LocalizedDisplayName -like "*Windows 10*22H2*x64*" } | 
  % { [PSCustomObject]@{ KB = [regex]::match($_.LocalizedDisplayName, 'KB(\d+)').value; Name = $_.LocalizedDisplayName; Description = $_.LocalizedDescription; Date = $_.DatePosted; } } |
  sort KB
  cd $SavedPath
}

function Get-CallingFileName {
  $cStack = @(Get-PSCallStack | ? { $_.ScriptName -and $_.ScriptName -notlike "*MBMod.psm1*" } )
  $cStack.ScriptName
}

function ScriptDir {
  #Only in local file
  if ($psise) { Split-Path $psise.CurrentFile.FullPath } else { $PSScriptRoot }
  #$global:GetScriptDir = { if ($psise) {Split-Path $psise.CurrentFile.FullPath} else {$PSScriptRoot} }
} 

function ListWindows {
  Get-Process | Where { $_.MainWindowTitle } | Select-Object ID, ProcessName, MainWindowTitle
}

function decom($pc) {
  CMinfo
  Get-CMDevice -Name $pc | Remove-CMDevice; C:
  $ws = ((Get-ItemProperty HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate -Name WUServer).WUServer) -split '//|:'
  $w = Get-WsusServer -Name $ws[2] -PortNumber $ws[3]
  $c = $w.SearchComputerTargets($pc)
  $c[0].Delete()
  Get-ADComputer $pc | Remove-ADObject -Recursive 
}

function Wsusinfo($server = 'drsopsmgr3') {
  $w = Get-WsusServer -Name drsopsmgr3 -PortNumber 8530
  $global:Wsus_comp = New-Object System.Collections.Generic.List[System.Object]
  $Wsus_comp = Get-WsusComputer -UpdateServer $w | ? { $_.ComputerRole -eq 'Workstation' }
  $global:WsusList = ($Wsus_comp.FullDomainName -replace '.dealers.aib.pri').ToUpper()
}

function CMinfo {
  Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1"
  Set-Location "DUB:\"
  $global:CM_comp = (Get-CMDevice | ? { $_.IsClient })
  $global:CMList = $CM_comp.name.ToUpper()
}

function SCCM-AppDetection {
  $AppName = "*Microsoft Visual C++ 2015-2022 Redistributable * - 14.36*"

  # Get OS-Native registry uninstall path:
  $RegPath = @("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
  If (${Env:ProgramFiles(x86)}) { $RegPath += "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall" }
  $FoundList = Get-ChildItem -Path $RegPath -Recurse | Get-ItemProperty | Where-Object { $_.DisplayName -like "*$AppName*" } 
  If ($FoundList) { $FoundList | % { [PSCustomObject]@{PC = $env:COMPUTERNAME; Name = $_.DisplayName; Ver = $_.DisplayVersion } } }

}

function Get-FileDetails($path) {
  $objShell = New-Object -ComObject Shell.Application 
  $objFolder = $objShell.namespace((Get-Item $path).DirectoryName) 

  foreach ($File in $objFolder.items()) {
    IF ($file.path -eq $path) {
      $FileMetaData = New-Object PSOBJECT 
      for ($a = 0 ; $a -le 266; $a++) {  
        if ($objFolder.getDetailsOf($File, $a)) { 
          $hash += @{$($objFolder.getDetailsOf($objFolder.items, $a)) = $($objFolder.getDetailsOf($File, $a)) }
          $FileMetaData | Add-Member $hash 
          $hash.clear()  
        } 
      }
    }
  }
  return $FileMetaData
}

function Clear-CalypsoOld {
  ADINFO
  $l = Ping-DealersPCs
  $pc = 'CCN7K4J-DUB'
  foreach ($pc in $l) {
    $dir = "\\$pc\C$\Program Files\CalypsoThickClient"
    $CalypsoDir = (gci $dir -Exclude Java, client)
    $CalypsoDir.Name
    #if ($CalypsoDir) { if (Test-Path $dir\Client) { $CalypsoDir.Fullname | % { $local=UncToLocal($_);  Run-Remote $pc "rd ""$local"" /s /q"  } } }
    $rest = (gci $dir).Name ;
    [PSCustomObject]@{ PC = $pc; CalypsoDir = $CalypsoDir.Name; Rest = $rest }
  }

}

function Get-Winver($pc) {
  # 10.0.19042 = 20H2      10.0.19044 = 20H2
  $build = (gwmi Win32_OperatingSystem -ComputerName $pc).Version 
  if ($build -eq '10.0.18362') { $ver = '19H1' } 
  if ($build -eq '10.0.18363') { $ver = '19H2' } 
  if ($build -eq '10.0.19041') { $ver = '20H1' }
  if ($build -eq '10.0.19042') { $ver = '20H2' } 
  if ($build -eq '10.0.19043') { $ver = '21H1' } 
  if ($build -eq '10.0.19044') { $ver = '21H2' }
  if ($build -eq '10.0.19045') { $ver = '22H2' }
  [PSCustomObject]@{ pc = $pc; ver = $ver; build = $build }
}

function GetUnc {
  [CmdletBinding()]param	( [Parameter(Mandatory = $True)] [string]$Path )
  $drive = (Get-Item $Path).PSDrive  #write $($script:MyInvocation.MyCommand.Definition) 
  $rest = Split-Path -Path "$Path" -NoQualifier
  $root = Get-PSDrive -Name $drive -ea 0 | select -ExpandProperty DisplayRoot
  if ($root) { $unc = Join-Path -Path $root -ChildPath $rest } #$drive.CurrentLocation
  if ($unc) { return $unc } else { return $path }
}

function DesktopPath {
  [Environment]::GetFolderPath("Desktop") + '\'
}

function sDate ($text) {
  if ($text) { "$text$(Get-Date -Format 'yyyy-MM-dd_HH-mm')" }
  else { "$(get-date -Format 'yyyy-MM-dd_HH-mm')" }
}

function MyTS ($timespan) {
  "{0:hh\:mm\:ss\.fff}" -f ([TimeSpan]$timespan)
  #((Get-BootTime $pc).up).Tostring("hh\:mm\:ss\.fff")
}

function IsEmail ($email) {
  [bool]($email -as [Net.Mail.MailAddress])
}

function Export-Xlsx ($obj, $path) {
  Export-Excel -NoNumberConversion Name -Path "$path" -InputObject $obj -TableName 'Table1' -TableStyle Medium7 -FreezeTopRow -BoldTopRow -AutoSize -CellStyleSB { param($workSheet)  $WorkSheet.Cells.Style.HorizontalAlignment = "Left" } 
}

function Export-Desktop ($obj, $text) {
  Export-Xlsx $obj "$(DesktopPath)$(sDate $text'_').xlsx"
}

function CombineObj ($ObjArray) { 
  $out = [PSCustomObject]@{ } 
  foreach ($o in $ObjArray) {
    foreach ($p in $o.psobject.Properties.name) {
      $name = $p;
      while ($name -in $out.PSObject.Properties.Name) { $name += '2' } 
      $out | Add-Member -MemberType NoteProperty -Name $name -Value $o.$p 
    }
  }
  $out
}

function Set-Proxy($val) {
  Set-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings' -name ProxyServer -Value 'webcorp.prd.aib.pri:8082'
  Set-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings' -name ProxyEnable -value $val
}

function Set-WorkWeekSchedule($ProgramName, $CollectionName, $time) {
  #Get-CMPackageDeployment -ProgramName $ProgramName | Select-Object PackageID -ExpandProperty AssignedSchedule 
  $a = 1..5 | % { New-CMSchedule -DayOfWeek $_ -Start (Get-Date -F "dd/MM/yy $time") }
  Get-CMDeployment -ProgramName $ProgramName -CollectionName $CollectionName | Set-CMPackageDeployment -StandardProgramName $ProgramName -Schedule $a  
}

function Get-RemoteReg ($PC, [Microsoft.Win32.RegistryHive] $HKEY, $Path, $name) {
  try {
    if ($HKEY -eq 'CurrentUser') {
      $HKEY = 'Users'
      $Path = "$((Get-RemoteReg $PC -HKEY Users).name | ? { $_ -like "S-1-5-21*"} | ? { $_ -notlike "*_Classes"})\$Path" 
      Write-Verbose "$PC\$HKEY\$path" 
    }
    $regBaseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($HKEY, $pc)
    $regKey = $regBaseKey.OpenSubKey($Path)
    if ($name) { if ($name -eq '(default)') { $name = "" }; $regKey.GetValue($name) } 
    else {
      if ($regkey) {
        ''; $regKey.Name;
        $regKey.GetSubKeyNames() | sort | % { [pscustomobject]@{Name = $_; Value = 'SubKey' } }
        $regkey.GetValueNames() | sort | % { [pscustomobject]@{Name = $_; Type = $regkey.GetValueKind($_); Value = $regkey.GetValue($_) } }
      } ; ''; 
    }
    $regkey.Close() 
  }
  catch { $false } 
}

function Set-RemoteReg ($PC, [Microsoft.Win32.RegistryHive] $HKEY, $Path, $name, $value, [Microsoft.Win32.RegistryValueKind] $kind) {
  try {
    if ($HKEY -eq 'CurrentUser') {
      $HKEY = 'Users'
      $Path = "$((Get-RemoteReg $PC -HKEY Users).name | ? { $_ -like "S-1-5-21*"} | ? { $_ -notlike "*_Classes"})\$Path" 
      Write-Verbose "$PC\$HKEY\$path" 
    }
    $regBaseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($HKEY, $pc)
    $regKey = $regBaseKey.OpenSubKey($Path, $true)
    $regKey.SetValue($name, $value)
    $regkey.Close()
  }
  catch { $false } 
}

function Set-RemoteRegRecursive ($PC, [Microsoft.Win32.RegistryHive] $HKEY, $Path, $name, $value, [Microsoft.Win32.RegistryValueKind] $kind) {
  try {
    if ($HKEY -eq 'CurrentUser') {
      $HKEY = 'Users'
      $Path = "$((Get-RemoteReg $PC -HKEY Users).name | ? { $_ -like "S-1-5-21*"} | ? { $_ -notlike "*_Classes"})\$Path" 
      Write-Verbose "$PC\$HKEY\$path" 
    }
    $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($HKEY, $PC)
    $path -split '\\' | % {
      $reg.CreateSubKey("$_", $true)
      $reg = $reg.OpenSubKey("$_", $true) 
    }
    $reg.SetValue($name, $value, $kind)
    $reg.Close()
  }
  catch { $false } 
}

function UncToLocal($path) {
  $path -replace '(?:.+)\\([a-z])\$\\', '$1:\'
}

function UncToLocal2($path) {
  $Drive = [System.IO.Path]::GetPathRoot($path)
  $Dumps = $path.Substring($Drive.Length)
  $Drive = $Drive.Substring($Drive.LastIndexOf('\') + 1).Replace('$', ':')
  $NTFSPath = "$Drive$Dumps"
}

function Run-Remote($Pc, $Cmd, $Timeout = 3, $CurrentDir = ’C:\temp’) {
  if (!(Aping $pc)) { 'Offline'; break }
  $opt = New-CimSessionOption -Protocol DCOM
  try {
    $s = New-CimSession -Computername $pc -SessionOption $opt -OperationTimeoutSec $timeout -ErrorAction Stop  
    Invoke-CimMethod Win32_Process -method Create @{CommandLine = "cmd /c $cmd"; CurrentDirectory = $CurrentDir } -CimSession $s
    Remove-CimSession $s 
  }
  catch { $false } 
}

# usage cmd : Run-Remote w10-mb "dir nosuchfile.txt > c:\temp\mm.txt 2>&1"
# usage ps  : Run-Remote W10-mb "powershell -command ""gci C:\Temp | Out-File C:\temp\aa_ll.txt"" "

function Run-Remote2($Pc, $Cmd, $Timeout = 3, $CurrentDir = ’C:\temp’) {
  $opt = New-CimSessionOption -Protocol DCOM
  try {
    $s = New-CimSession -Computername $pc -SessionOption $opt -OperationTimeoutSec $timeout -ErrorAction Stop  
    Invoke-CimMethod Win32_Process -method Create @{CommandLine = $cmd; CurrentDirectory = $CurrentDir } -CimSession $s
    Remove-CimSession $s 
  }
  catch { $false } 
}

<#
AIB MENU
$c = "msiexec.exe /qn /i ""C:\Temp\Rocket Passport To PC Host\Passport.msi"" /quiet /qn LICENSE=""DRKV-FG92-1COQ-KB7P"" ALLUSERS=2 USERDATADIR=""C:\Program Files (x86)\PASSPORT\"""
$pc = "6NS9MM2-DUB"
Run-Remote $pc $c
#>

function Run-RemoteCred($Pc, $Cmd, $Timeout = 3, $CurrentDir = ’C:\temp’, $cred) {
  $opt = New-CimSessionOption -Protocol DCOM
  try {
    $s = New-CimSession -Computername $pc -SessionOption $opt -OperationTimeoutSec $timeout -ErrorAction Stop -Credential $cred
    Invoke-CimMethod Win32_Process -method Create @{CommandLine = "cmd /c $cmd"; CurrentDirectory = $CurrentDir } -CimSession $s
    Remove-CimSession $s 
  }
  catch { $false } 
}

function Run-Remote_WMIold($pc, $cmd) {
 ([WMICLASS]"\\$pc\ROOT\CIMV2:win32_process").Create($cmd).ProcessId
}

function Check-WMI($pc, $timeout = 3) {
  $opt = New-CimSessionOption -Protocol DCOM
  try {
    $s = New-CimSession -Computername $pc -SessionOption $opt -OperationTimeoutSec $timeout -ErrorAction Stop 
    $t = (get-date) - (gcim Win32_OperatingSystem -CimSession $s -ErrorAction SilentlyContinue).LastBootUpTime
    Remove-CimSession $s 
    [bool]$t
  }
  catch { $false } 
}

function Get-MissingDrivers($pc) {
  #For formatting:
  $result = @{Expression = { $_.Name }; Label = "Device Name" },
  @{Expression = { $_.ConfigManagerErrorCode } ; Label = "Status Code" }

  #Checks for devices whose ConfigManagerErrorCode value is greater than 0, i.e has a problem device.
  Get-WmiObject -Class Win32_PnpEntity -ComputerName $pc -Namespace Root\CIMV2 | Where-Object { $_.ConfigManagerErrorCode -gt 0 } | select name, ConfigManagerErrorCode #| Format-Table $result -AutoSize
}

function MyProgress ($text, $maxcount) {
  #rv ii -ErrorAction SilentlyContinue
  If (-not(Test-Path Variable:\ii)) { $global:ii = 0 }
  $global:ii++
  If ($global:ii -gt $maxcount) { $global:ii = 0 } 
  $perc = [math]::Round($ii / $maxcount * 100, 1);
  Write-Progress $text "Complete : $perc %" -perc $perc
}

function Menu ($Title, [array]$opt) {
  "$Title"
  '-' * 20
  for ($i = 1; $i -lt $opt.count + 1; $i++) { 
    if ($opt[$i - 1] -eq 'back to search') 
    { "[0] $($opt[$i-1])" } 
    else
    { "[$i] $($opt[$i-1])" }
  }
}

function hist ($o) {
  #load from file
  if (-not $o) { $o = 'Test entry ..' }
  if (-not(Test-Path variable:global:hist)) { [System.Collections.Generic.List[object]] $global:hist = @() }
  $global:hist.Insert(0, $o)
  $global:hist = $global:hist | select -first 20
  #save to file
}

function hl ($text, $word, $fc = 14, $bc, $nonewline) {
  $text = ($text | Out-String).Trim()
  $s = $text -split ([regex]::Escape($word))
  Write-Host $s[0] -NoNewline
  for ($i = 1; $i -lt $s.count; $i++) {  
    $ex = "Write-Host $word -NoNewline -ForegroundColor $fc "
    if ($bc) { $ex += "-BackgroundColor $bc" } 
    iex $ex
    Write-Host $s[$i] -NoNewline
  }
  if (!$nonewline) { Write-Host }
}

[Management.Automation.WildcardPattern]::Escape('test[1].txt (foo)')
[regex]::Escape("test[1].txt (foo)")

Function WinTitle($Title) {
  $host.ui.RawUI.WindowTitle = $Title
}

function Get-UserProfile($pc) {
  $opt = New-CimSessionOption -Protocol DCOM
  $s = New-CimSession -Computername $PC -SessionOption $opt -ErrorAction Stop
 (Get-CimInstance -Class Win32_UserProfile -CimSession $S).LocalPath | % { $_.split('\')[-1] } | ? { $_ -match "\d" }
}

function Remove-UserProfile($PC, $user) {
  "$pc - $user"
  $opt = New-CimSessionOption -Protocol DCOM
  $s = New-CimSession -Computername $PC -SessionOption $opt -ErrorAction Stop
  $all = (Get-CimInstance -Class Win32_UserProfile -CimSession $S).LocalPath | % { $_.split('\')[-1] } | ? { $_ -match "\d" }
  if ($user) { Get-CimInstance -Class Win32_UserProfile -CimSession $S | Where-Object { $_.LocalPath.split('\')[-1] -eq $user } | Remove-CimInstance }
  else { $all }
  Remove-CimSession $s
}

function Remove-AllUsersProfile($PC) {
  "$pc - List of user profiles"
  $opt = New-CimSessionOption -Protocol DCOM
  $s = New-CimSession -Computername $PC -SessionOption $opt -ErrorAction Stop
  $all = (Get-CimInstance -Class Win32_UserProfile -CimSession $S).LocalPath | % { $_.split('\')[-1] } | ? { $_ -match "\d" }
  $all
  pause
  Get-CimInstance -Class Win32_UserProfile -CimSession $S | Where-Object { $_.LocalPath.split('\')[-1] -in $all } | Remove-CimInstance -Verbose
  Remove-CimSession $s
}

function ADinfo {
  #Write-Debug "Updating AD from servers"
  Init
  If (-not(Get-module ImportExcel)) { Import-Module "$ModuleDir\ImportExcel\7.4.1\ImportExcel.psd1" -Global -WA SilentlyContinue } 
  $null = Get-DealersUsers 
  $null = Get-DealersPCs
  Export-Xlsx $ADu $upath 
  Export-Xlsx $ADc $cpath
}

function Get-ADinfo {
  if (Test-Path $upath) {
    $time = gpv $upath -Name LastWriteTime
    if (((get-date) - $time).TotalHours -lt 2) {
      "-- Loaded from file --"; $global:ADu = Import-Excel $upath; $global:ADc = Import-Excel $cpath
    }
    else { "-- Loaded from AD --"; ADinfo }
  }
  else { "-- Loaded from AD --"; ADinfo } 

}

function Get-DealersUsers {
  $global:ADu = New-Object System.Collections.Generic.List[System.Object]
  $tempU = New-Object System.Collections.Generic.List[System.Object]
  $prop = @('msDS-UserPasswordExpiryTimeComputed', 'Name', 'DisplayName', 'Description', 'Office', 'mail', 'LastBadPasswordAttempt', 'BadPwdCount', 'LockedOut', 'pwdLastSet', 'proxyAddresses')
  #$global:ADu = Get-ADUser -Filter * -Properties $prop | ? { $_.name -match '^\d{5}$' } 
  $tempU.AddRange( (Get-ADUser -Filter * -Properties $prop ) )  # | ? { $_.name -match '^\d{5}$' }
  $tempU | % {
    $val = if ($_.'msDS-UserPasswordExpiryTimeComputed' -eq '9223372036854775807') { 'Password Never Expired' }
    else { Get-Date ([DateTime]::FromFileTime([Int64]::Parse($_.'msDS-UserPasswordExpiryTimeComputed'))) } # -Format "dd/MM/yyyy HH:mm:ss"   ([datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")) 
    $_ | Add-Member -MemberType NoteProperty -Name 'ExpiryDate' -Value $val -Force
    $_ | Add-Member -MemberType NoteProperty -Name 'LastPwdSet' -Value (Get-Date ([DateTime]::FromFileTime([Int64]::Parse($_.pwdLastSet)))) -Force
    # not needed ? if ($_.pwdLastSet) { $_.pwdLastSet  }
  } 
  $ex = 'msDS-UserPasswordExpiryTimeComputed', 'pwdLastSet', 'WriteDebugStream', 'WriteErrorStream', 'WriteInformationStream', 'WriteVerboseStream', 'WriteWarningStream', 'PropertyNames', 'AddedProperties', 'RemovedProperties', 'ModifiedProperties', 'PropertyCount'
  $ADu.AddRange( ($tempU | select * -ExcludeProperty $ex) )
  Remove-Variable TempU
  $r = 'St Helens', 'London', '1st Floor,', ' 1 Undershaft', 'Old Jewry' -join '|'
  $ADu | % { $_.office = ($_.office -replace $r).trim() }
}

function Get-DealersPCs {
  # $s = ([adsisearcher]"(&(objectCategory=computer)(!(operatingsystem=*Server*))((operatingsystem=*)))").FindAll().Properties.cn | ? { $_ -ne 'DRSVCENTRE'}

  $global:ADc = New-Object System.Collections.Generic.List[System.Object]
  $TempC = New-Object System.Collections.Generic.List[System.Object]
  $TempC.AddRange( ((Get-ADComputer -Filter { OperatingSystem -NotLike "*server*" } -prop description, location) | ? { $_.name -ne 'DRSVCENTRE' }) )
  $ex = 'PropertyNames', 'AddedProperties', 'RemovedProperties', 'ModifiedProperties', 'PropertyCount'
  $ADc.AddRange( ($TempC | select * -ExcludeProperty $ex) )
  Remove-Variable TempC
}

function Ping-DealersPCs {
  APingN($adc.name)
}

function LockoutStatusJob ($user) {
  Remove-Job -Name 'LockoutStatus' -ea SilentlyContinue
  $sc = { iex ${using:function:ImportMe}.Ast.Extent.Text; ImportMe;
    LockoutStatus $using:user
  }
  Start-Job -Name 'LockoutStatus' -ScriptBlock $sc
}

function LockoutStatus ($user) {
  $DCs = New-Object System.Collections.Generic.List[System.Object]
  $DCs.AddRange( (Get-ADDomainController -Filter * | ? { $_.name -ne 'DRSGAMMADC1' } ) )
  $DCs.AddRange( (Get-ADDomainController -Filter * -Server prd.aib.pri | Select -First 6) )
  $online = APing($DCs.hostname)
  Foreach ($DC in $online) {
    $t = Get-ADUser -Identity $user -Server $DC.Name -Properties AccountLockoutTime, LastBadPasswordAttempt, BadPwdCount, LockedOut, pwdLastSet, msDS-UserPasswordExpiryTimeComputed
    if ($t) {
      Add-Member -InputObject $t -MemberType NoteProperty -Name DC -Value $DC.Name -Force
      Add-Member -InputObject $t -MemberType NoteProperty -Name LastPwdSet -Value (Get-Date ([DateTime]::FromFileTime([Int64]::Parse($t.pwdLastSet)))) -Force
      Add-Member -InputObject $t -MemberType NoteProperty -Name ExpiryTime -Value (Get-Date ([DateTime]::FromFileTime([Int64]::Parse($t.'msDS-UserPasswordExpiryTimeComputed')))) -Force 
    }
    else { $dc.name }
    $t | Select DC, Name, Enabled, LockedOut, @{N = 'LastBad'; E = { $_.LastBadPasswordAttempt } }, @{N = 'BadCount'; E = { $_.BadPwdCount } }, LastPwdSet, ExpiryTime
  }  
}

function Set-Console($title, $width, $height) {
  if ($title) { $host.UI.RawUI.WindowTitle = $Title }
  if ($width) { [console]::WindowWidth = $width; [console]::BufferWidth = [console]::WindowWidth }
  if ($height) { [console]::WindowHeight = $height }
}

function SetWinTitle($p, $text) {
  if ("Win32Api" -as [type]) {} else {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
  
public static class Win32Api
{
    [DllImport("User32.dll", EntryPoint = "SetWindowText")]
    public static extern int SetWindowText(IntPtr hWnd, string text);
}
"@
  }
  # How to use 
  #$p = Start-Process -FilePath "notepad.exe" -PassThru
  #$p.WaitForInputIdle() | out-null #only GUI
  [Win32Api]::SetWindowText($p.MainWindowHandle, $text)  
}

Function Execute-Command ($commandTitle, $commandPath, $commandArguments) {
  #HOWTO: $DisableACMonitorTimeOut = Execute-Command -commandTitle "Disable Monitor Timeout" -commandPath "C:\Windows\System32\powercfg.exe" -commandArguments " -x monitor-timeout-ac 0"

  $Psexec = (Get-Module invokepsexec).ModuleBase + '\PsExec.exe'
  $pinfo = New-Object System.Diagnostics.ProcessStartInfo
  $pinfo.FileName = $Psexec #$commandPath
  $pinfo.RedirectStandardError = $true
  $pinfo.RedirectStandardOutput = $true
  $pinfo.UseShellExecute = $false
  $pinfo.Arguments = "\\$pc cmd" #$commandArguments
  $p = New-Object System.Diagnostics.Process
  $p.StartInfo = $pinfo
  $p.Start() 
  #$p.WaitForExit()
  [pscustomobject]@{
    commandTitle = $commandTitle
    stdout       = $p.StandardOutput.ReadToEnd()
    stderr       = $p.StandardError.ReadToEnd()
    ExitCode     = $p.ExitCode
  }


  $sb_new = { 
    iex ${using:function:ImportMe}.Ast.Extent.Text; ImportMe; Test-Modules
    $Psexec = (Get-Module invokepsexec).ModuleBase + '\PsExec.exe' 

  }
  Start-Job 

}

function RemoteCmd($pc) {
  $Psexec = (Get-Module invokepsexec).ModuleBase + '\PsExec.exe'   # & $psexec \\$pc cmd.exe  # same window
  $proc = start "$psexec" "\\$pc cmd" -PassThru                       # Invoke-PsExec $pc -Command 'hostname'
  sleep -m 1500
  SetWinTitle $proc "cmd on $pc"
}

function New-PSWin($in) {
  start powershell
}

function New-PSWin-Alert($in) {
  #do not use security alerts
  invoke-expression "cmd /c start powershell -NoExit -Command {  Get-date;                      `
     $($function:ImportMe.Ast.Extent.Text); ImportMe; Set-Console 'Title' 80 25; Test-Modules;  `
     cd `$ScriptPath; Get-ADinfo; `$in = '$($in | ConvertTo-Json)' | ConvertFrom-Json           `
}"                                #to do : in - user data in executed expresion
}

Function Check-User ($user) {
  #ADinfo
  $u = $ADu | ? { $_.Name -eq $user }
  $j = LockoutStatusJob $user      # pwdLastSet
  "`n" * 2
  ($u | select Name, DisplayName, description, office, LastPwdSet, ExpiryDate | fl | Out-String ).Trim() 
  #Write-host "`nChecking where the user last logged in, found computers : " -NoNewline
  $l = Get-LoggedUsers
  [array]$out = $LoggedUsers | ? { $_.username -eq $user } | select -Unique 
  Write-host '' #$out.Count
  if ($out) {
    if ( $u.Office -and $u.Office -notin $out.computer + ''  ) { $out += (Logged-User $u.Office) } #-and $u.office -in $adc.name -and (APingN $u.offlice)
    $out | % { $_ | Add-Member -MemberType NoteProperty -Name UpTime -Value (Get-BootTimeF $_.computer) -Force }
    $out | % { $_ | Add-Member -MemberType NoteProperty -Name LoggedNow -Value (Logged-User $_.computer).USERNAME -Force }
    $out | % { $x = $_.LoggedNow; $_ | Add-Member -MemberType NoteProperty -Name LoggedNowDN -Value ($ADu | ? { $_.Name -eq $x }).DisplayName }
    ($out | select Computer, Description, UpTime, 'LOGON TIME', LoggedNow, LoggedNowDN | ft | Out-String ).Trim() #-HideTableHeaders
  } 
  
  $pc = ($out | ? { $_.LoggedNow -eq $user }).Computer
  if (-not $pc) { $pc = $u.Office }
  if (-not $pc) { $pc = $out.Computer | select -First 1 }

  if ($pc) { $pc | Set-Clipboard; "`n'$pc' has been copied to the clipboard" }
  ''
  Menu "Choose option" @('Show Lockout Status', 'Show users AD groups', "Go to $pc", 'New console window', 'Back')
  '' 
  $inp = Read-Host "[1-5] "
  switch ($inp) {
    '1' { Receive-Job -Name 'LockoutStatus' -Wait | select * -ExcludeProperty RunspaceId, PSSourceJobInstanceId | ft }
    '2' { "`n$($u.Name) is a member of:"; Get-UserGroup $u.Name; '' }
    '3' { Check-PC $pc }       # "Set-ADAccountPassword $u -Reset -NewPassword (ConvertTo-SecureString -AsPlainText 'p@ssw0rd' -Force) " }
    '4' { New-PSWin $user }
    '0' { "back to search" }
    Default { "back to search" }
  }
}

function Check-PC($pc) {
  # take a look something is taking lots of times sometimes
  $p = $ADc | ? { $_.Name -eq $pc }
  $on = APing($pc)
  $l = Get-LoggedUsers; [array]$LLast = $LoggedUsers | ? { $_.Computer -eq $pc }
  "`n" * 2
    ($p | select Name, description, DNSHostName | fl | Out-String ).Trim()  
  if ($on) { 
    $uptime = Get-BootTimeF $pc
    $LNow = Logged-User $pc    
    "Online      : $($on.Address)" 
    "Up Time     : $uptime"
    "Logged User : $((isLogged $pc).user) $($LNow.USERNAME)  $($LNow.DisplayName)  $($LNow.'LOGON TIME')  $($LNow.SESSIONNAME)"
        
  }
  else { "Offline !! " }; ''

  $Opt = @( "Open C: - \\$pc\c$",
    "Open comand prompt on $pc",
    "DameWare $pc",
    "Remote Desktop $pc"
    'Show computer AD groups',
    'Wake On Lan',
    'Ping',
    'Restart',
    "Computer Management $pc"
    'back to search' );

  Menu "Choose option" $Opt 
  '' 
  $inp = Read-Host "[1-$($Opt.count)] "
  switch ($inp) {
    '1' { ii "\\$pc\c$" }
    '2' { RemoteCmd $pc }
    '3' { New-DameWare $pc }
    '4' { New-MSTSC $pc }
    '5' { "`n$pc is a member of:"; Get-PCgroup $pc; '' }
    '6' { WOL $pc; New-PingWindow($pc) }
    '7' { New-PingWindow($pc) }
    '8' { Restart-Computer $pc -Force; New-PingWindow($pc) }
    '9' { compmgmt.msc -a /computer=$pc }
    '0' { "back to search" }
    Default { "back to search" }
  }
}

function New-DameWare($pc) {
  $dw = "C:\H\_Apps\DameWare Mini Remote Control\DWRCC.exe"
  $cmd = "-m:$pc -a:1" # -h -c"
  #iex  "&'$dw' $cmd"
  start "$dw" "$cmd"     
}

function New-MSTSC($PCs) {
  $PCs | % { mstsc /v:$_ }
}

function New-PingWindow($ip) {
  start-process cmd -ArgumentList "/C", "mode con:cols=55 lines=10 && title Ping $ip && powershell -command ""&{(get-host).ui.rawui.buffersize=@{width=55;height=200};}"" && ping $ip -t"
}

function Get-ADRealUsers {
  Get-ADUser -Filter { Surname -like "*" -and memberof -like '*' } -prop name, givenname, surname `
  | select name, givenname, surname # | Export-Excel -Path C:\Users\dsk_58691\Desktop\usr.xlsx
}

function Get-GraphicDrivers($pc) {
  Get-WmiObject Win32_VideoController -ComputerName $pc | ForEach-Object {
    [PSCustomObject]@{
      ComputerName  = $_.SystemName
      Description   = $_.Description -join ', '
      DriverDate    = [DateTime]::ParseExact($_.DriverDate -replace '000000.000000-000', 'yyyyMMdd', $culture).ToString('yyyy-MM-dd')
      DriverVersion = $_.DriverVersion
      PNPID         = $_.PNPDeviceID
      # desc          = ($ad | ? { $_.name -eq $pc }).description
    }
  }
}

function Get-ExpiringUsers ($days) {
  $WarnDate = (get-date).adddays($days)
  $users = @()  # init array
  $users = Get-ADUser -filter { Enabled -eq $True -and PasswordNeverExpires -eq $False -and PasswordLastSet -gt 0 -and Name -notlike "*$*" } `
    –Properties Name, DisplayName, msDS-UserPasswordExpiryTimeComputed, EmailAddress, UserPrincipalName `
  | Select-Object -Property Name, Displayname, @{Name = "ExpiryDate"; Expression = { [datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed") } }, EmailAddress, UserPrincipalName `
  | Where { $_.ExpiryDate -gt (Get-Date) -and $_.ExpiryDate -le $WarnDate } `
  | Sort-Object ExpiryDate   #" $($users.count) users with a password expiring between $((Get-Date).ToShortDateString()) and $($WarnDate.ToShortDateString()) "
  #$users | Out-GridView -PassThru -Title "Select users, use CTRL or SHIFT to select many" | SendEmailByOutlook 
  $users
}

function Get-LoggedUsers {
  
  $sb_new = { 
    iex ${using:function:ImportMe}.Ast.Extent.Text; ImportMe; Test-Modules
    Init; Get-ADinfo
    $log = (APingN($ADc.name)) | Logged-User
    $file = "$ModulePath\db\$(sDate 'Logged').xlsx" 
    Export-Xlsx -obj $log -path $file
    $file
  }

  $lpath = "$ModulePath\db\" + "logged*.xlsx"
  $files = gci $lpath | sort LastWriteTime -Descending 
  if ($files) {
    if (((get-date) - $files[0].LastWriteTime).TotalHours -lt 1) {
      #"-- Loaded from file --" + $files[0].LastWriteTime.ToString("yyyy/MM/dd hh:mm");     
    }
    else {
      #"-- Need update --"
      Remove-Job -Name 'LoggedUserJob' -ErrorAction SilentlyContinue
      $job = Start-Job -Name 'LoggedUserJob' -ScriptBlock $sb_new 
    }
    if ($files.Count -gt 8) {
      $zip = gci $lpath | sort LastWriteTime -Descending | select -Skip 1
      $temp = New-Object System.Collections.Generic.List[System.Object]
      $zip | % { $temp.AddRange( (Import-Excel $_) ) }
      $temp = $temp  | ? { $_.'LOGON TIME' } | Sort-Object Computer -Unique -Descending | Sort-Object dt -Descending
      Remove-Item $files
      $temp | Export-Excel -Path "$ModulePath\db\Logged.xlsx" -TableName 'Table1' -TableStyle Medium7 -FreezeTopRow -BoldTopRow -AutoSize
    }
  }
  else {
    "-- No files --, updating, please wait a minute"
    $job = Start-Job -Name LoggedUserJob -ScriptBlock $sb_new | wait-job 
  } 
 
  $sb_import = {
    $global:LoggedUsers = New-Object System.Collections.Generic.List[System.Object]
    $global:LoggedLast = New-Object System.Collections.Generic.List[System.Object]
    $files = gci $lpath | sort LastWriteTime -Descending
    $LoggedLast.AddRange( (Import-Excel $files[0].FullName) )
    $files | % { $LoggedUsers.AddRange( (Import-Excel $_) ) }
    $LoggedUsers = ($LoggedUsers | ? { $_.Username -ne 'NONE' -and $_.displayName }) ### !!!!
  }

  # if LoggedUsers not exist
  & $sb_import
}

function Logged-User {
  [CmdletBinding()]Param([Parameter(ValueFromPipeline)]$pc)
  begin { if (-not (Test-Path variable:adu) ) { ADinfo } }
  process {
    if ($pc -eq "") { $pc = $env:COMPUTERNAME }
    $o = [PScustomObject]@{ Computer = $pc; Description = ($Adc | ? { $_.name -eq $pc }).Description; 
      USERNAME = ''; DisplayName = ''; SESSIONNAME = ''; ID = ''; STATE = ''; 'IDLE TIME' = ''; 'LOGON TIME' = '';
      dt = (get-date -Format G)    
    }
    if (APing $pc) {
      try {
        $temp = (query user /server:$pc 2>&1)  
        If ($temp) {
          # If ($temp -split '`n' -eq 'No User exists for *') {$temp = $null; $user = $false}
          $r = $temp -replace '\s{2,}', ',' | ConvertFrom-Csv
          $r.psobject.Properties.name | % { $o.$_ = $r.$_ }
          $o.DisplayName = ($adu | ? { $_.name -eq $r.USERNAME }).DisplayName 
        }
      }
      catch { $o.USERNAME = 'NONE' }
    }
    else { $o.USERNAME = 'OFFLINE' }
    $o
  }
}

function Logged-User2 {
  [CmdletBinding()]Param([Parameter(ValueFromPipeline)]$pc)
  process {
    
    if ($pc -eq "") { $pc = $env:COMPUTERNAME }
    $o = [PScustomObject]@{ Computer = $pc; Description = ($Adc | ? { $_.name -eq $pc }).Description; 
      USERNAME = ''; DisplayName = ''; SESSIONNAME = ''; ID = ''; STATE = ''; 'IDLE TIME' = ''; 'LOGON TIME' = '';
      dt = (get-date -Format G)    
    }
    if (APing $pc) {
      try {
        $temp = (query user /server:$pc 2>&1)  
        If ($temp) {
          # If ($temp -split '`n' -eq 'No User exists for *') {$temp = $null; $user = $false}
          $r = $temp -replace '\s{2,}', ',' | ConvertFrom-Csv
          $r.psobject.Properties.name | % { $o.$_ = $r.$_ }
          $o.DisplayName = ($adu | ? { $_.name -eq $r.USERNAME }).DisplayName 
        }
      }
      catch { $o.USERNAME = 'NONE' }
    }
    else { $o.USERNAME = 'OFFLINE' }
    $o
  }
}



function isLogged($pc = "$env:COMPUTERNAME") {
  $i = 0; $user = $null; $r = $null
  if (APing($pc)) {
    try {
      $temp = (query user /server:$pc 2>&1)  
      If ($temp -split '`n' -eq 'No User exists for *') { $temp = $null; $user = $false }
      If ($temp) { 
        $r = $temp -replace '\s{2,}', ',' | ConvertFrom-Csv 
        If ($r.USERNAME[0] -eq '>') { $r.USERNAME = $r.username.Substring(1) }
        $User = $r.USERNAME 
      }
    }
    catch { $user = 'error' }
  }
  else { $user = 'pcoff' }
  return [PSCustomObject]@{ PC = $pc; User = $user }
} 

function Get-BootTime ($pc) {
  $opt = New-CimSessionOption -Protocol DCOM
  try {
    $s = New-CimSession -Computername $pc -SessionOption $opt -OperationTimeoutSec 3 -ErrorAction Stop
    $t = (get-date) - (gcim Win32_OperatingSystem -CimSession $s -ErrorAction SilentlyContinue).LastBootUpTime
    Remove-CimSession $s 
  }
  catch { $t = 0 }
    
  [PScustomObject]@{ PC = $pc; up = $t; }
}

function Get-BootTimeF ($pc) {
  (Get-BootTime $pc).up.tostring("dd\.hh\:mm\:ss")
}

function Get-UnusedCN($uptime = 72) {
  Write-Progress "Getting list of unused computers" "..." -perc 0
  $l = (Get-ADComputer -Filter { OperatingSystem -NotLike "*server*" }).name # Write-verbose "Getting list of computers from AD where OperatingSystem is not like *server*" -and Name -like "*-DUB"
  Write-Progress "Getting list of unused computers" "Ping.." -perc 25
  $on = ( APing($l) ).name #Write-verbose "Ping list of computers"
  Write-Progress "Getting list of unused computers" "Logged.." -perc 50
  $listLog = $on | % { isLogged $_ } #Write-verbose "Checking users logged $($on.Count) online"
  $nouser = $listLog | ? { $_.User -eq $false } #Write-verbose "Get computers without user logged $($listLog.Count) logon "
  Write-Progress "Getting list of unused computers" "BootTime.." -perc 75
  $times = $nouser.pc | % { Get-BootTime $_ } #Write-verbose "Get boot times TotalHours > $($uptime). $($nouser.Count) unused computers"
  $togo = $times | ? { $_.up.TotalHours -gt $uptime } | % { $_ | Add-Member -MemberType NoteProperty -Name Desc -Value (Get-ADComputer $_.pc -Properties Description).Description -PassThru -force }
  Write-Progress "Done.." -Completed
  $global:togo = $togo
  $global:togo
}

function Restart-Unused {
  $togo = Get-UnusedCN
  Write-Host "`nFollowing computers will be resarted now : `n"
  $togo | ft
  pause
  $script:restarted = @() #Start-Transcript "$ScriptPath\RebootLog.txt" -Append 
  $togo | % { if ( (isLogged $_.pc).user -eq $false ) { $_; Restart-Computer $_.pc -Force -Verbose -ErrorAction SilentlyContinue; $restarted += , $_.pc } }
  $restarted | Out-File "$ScriptPath\$(sdate RestartLog).txt" -Append #Stop-Transcript
}

function Check-Logs ($pc, $LastHours) {
  # calculate start time (one hour before now)
  $Start = (Get-Date) - (New-Timespan -Hours $LastHours) 
  # Getting all event logs
  Get-EventLog -AsString -ComputerName $pc |
  ForEach-Object {
    # write status info
    Write-Progress -Activity "Checking Eventlogs on \\$pc" -Status $_
    # get event entries and add the name of the log this came from
    Get-EventLog -LogName $_ -EntryType Error, Warning -After $Start -ComputerName $pc -ErrorAction SilentlyContinue |
    Add-Member NoteProperty EventLog $_ -PassThru      
  } |
  # sort descending
  Sort-Object -Property TimeGenerated -Descending |
  # select the properties for the report
  Select-Object EventLog, TimeGenerated, EntryType, Source, Message | 
  # output into grid view window
  Out-GridView -Title "All Errors & Warnings from \\$pc"
}

function Get-UpdatesRemotly ($pc) {
  if ($pc -eq "") { $pc = $env:COMPUTERNAME }
  Try {
    $session = [activator]::CreateInstance([type]::GetTypeFromProgID(“Microsoft.Update.Session”, $pc))
    $searcher = $session.CreateUpdateSearcher()
    $totalupdates = $searcher.GetTotalHistoryCount()
    $all = $searcher.QueryHistory(0, $totalupdates)
  }
  catch { "Cannot connect : $pc" }
    
  $j = 0
  $Out = @()
  Foreach ($update in $all) {
    Write-Progress "Processing updates on: $pc" "Complete : $j of $($all.count)" -perc (($j / $all.count) * 100); $j++
    if ($update.operation -eq 1 -and $update.resultcode -eq 2) {
      $Out += [pscustomobject]@{
        'ComputerName'        = $pc
        'UpdateDateTime'      = $update.date
        'KB'                  = [regex]::match($update.Title, 'KB(\d+)')
        'UpdateTitle'         = $update.title
        'SupportUrl'          = $update.SupportUrl
        'UpdateDescription'   = $update.Description
        'UpdateId'            = $update.UpdateIdentity.UpdateId
        'RevisionNumber'      = $update.UpdateIdentity.RevisionNumber
        'Operation'           = $update.Operation
        'ResultCode'          = $update.ResultCode
        'HResult'             = $update.HResult
        'UnmappedResultCode'  = $update.UnmappedResultCode
        'ClientApplicationID' = $update.ClientApplicationID
        'ServerSelection'     = $update.ServerSelection
        'ServiceID'           = $update.ServiceID
        'UninstallationNotes' = $update.UninstallationNotes
        #'UninstallationSteps' = $update.UninstallationSteps
      }
    }
  }
  $Out 
}

function ExpStr($string) {
  $ExecutionContext.InvokeCommand.ExpandString($string)
}

function Check-Paths ($PCs, $paths) {
  #Paths should be ' ' surrounded if $pc var needed eg '\\$pc\C$\Program Files (x86)\SAP'
  $on = APingN($PCs)
  [System.Collections.ArrayList]$all = @()
  $cnt = ($on.count * $paths.count)
  foreach ($pc in $on) { 
    foreach ($path in $paths) {
      $p = ExpStr($path)
      MyProgress "$pc - $path" $cnt
      $o = [pscustomobject]@{ PC = $pc; Path = $p; Exist = Test-Path -Path $p; }
      $o
      [void]$all.Add($o) 
    }
  } 
  #Export-Xlsx $all "$(DesktopPath)Folder.xlsx"
}

function Map-Adrive {
  <# .Example
      Map-Adrive Z \\server\folder #>    
  [CmdletBinding()]
  param(
    [string]$driveletter,
    [string]$path,
    [switch]$persistent
  )
  process {
    $nwrk = new-object -com Wscript.Network
    Write-Verbose "Mapping $($driveletter+':') to $path and persist=$persistent"
    try {
      $nwrk.MapNetworkDrive($($driveletter + ':'), $path)     
      Write-Verbose "Mapping successful."
    }
    catch {
      Write-Verbose "Mapping failed!"
    }
  }
}

function WOL {
  [CmdletBinding()] Param(
    $CmpName = $Null,
    $CollId = $Null, #"SMS00001"
    $SiteServer = "drscmsrv2.dealers.aib.pri"
  )
 
  Write-Verbose "CmpName = $CmpName"
  Write-Verbose "CollId  = $CollID"
  Write-Verbose "SiteServer = $SiteServer"

  if (!$CmpName -and !$CollId) { Write-Warning "Please provide ComputerName or CollectionID to WOL" ; break }
  if (!$CmpName -and $CollId -eq "SMS00001") {
    Write-Warning "Seems wrong to wake every single computer in the environment, refusing to perform." ; break  
  }
 
  $SiteCode = (Get-WmiObject -ComputerName "$SiteServer" -Namespace root\sms -Query 'SELECT SiteCode FROM SMS_ProviderLocation').SiteCode
 
  if ($CmpName) {
    $ResourceID = (Get-WmiObject  -ComputerName "$SiteServer" -Namespace "Root\SMS\Site_$($SiteCode)" -Query "Select ResourceID from SMS_R_System Where NetBiosName = '$($CmpName)'").ResourceID
    if ($ResourceID) { $CmpName = @($ResourceID) }
  }
 
  $WMIConnection = [WMICLASS]"\\$SiteServer\Root\SMS\Site_$($SiteCode):SMS_SleepServer"
  $Params = $WMIConnection.psbase.GetMethodParameters("MachinesToWakeup")
  $Params.MachineIDs = $CmpName
  $Params.CollectionID = $CollId
  $return = $WMIConnection.psbase.InvokeMethod("MachinesToWakeup", $Params, $Null) 
 
  if (!$return) {
    Write-Host "No machines are online to wake up selected devices" 
  }
  if ($return.numsleepers -ge 1) {
    Write-Host "The resource selected are scheduled to wake-up as soon as possible" 
  } 
}

Function SendEmailByOutlook {
  [CmdletBinding()]Param([Parameter(ValueFromPipeline)]$user)
  process {
    $currPath = if ($psISE) { Split-Path $psISE.CurrentFile.FullPath } else { $PSScriptRoot } 
    $file = Resolve-Path ($currPath + "\ExpiryReminder.msg")
    $u = $user
    $name = ($u.Displayname -split ' ')[0]
    $now = Get-Date
    $exp = $u.ExpiryDate - $now.Date
    $expDays = $exp.Days
    $ol = New-Object -ComObject outlook.application -Verbose:$false
    $msg = $ol.CreateItemFromTemplate($file) 
    $msg.To = if ($u.EmailAddress) { $u.EmailAddress } else { $u.UserPrincipalName }
    $sDay = if ($u.ExpiryDate.Date -eq $now.Date) { "today !" } else {
      if ($u.ExpiryDate.Date -eq $now.Date.AddDays(1)) { "tomorrow" } 
      else { if ($expDays -gt 1) { "in $expDays days" } else { "in $expDays day" } }
    }
    $s1 = "Your password is due to expire $sDay" 
    $msg.Subject = "Your password will expire $sDay"
    $msg.HTMLbody = $msg.HTMLbody.Replace("Hi Folks", "Hi $name") 
    $msg.HTMLbody = $msg.HTMLbody.Replace("is due to expire in the next 7 Days", "will expire $sDay")

    $msg.Display()

    1..2 | % {
      $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($msg)
      $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ol)
    }
  }
}

function Convert-ToUnc ($localPath, $pc) {    
  [void]( $l = $localPath -replace '^(.):', "\\$pc\`$1$" )
  return $l
}

Function CopyWin {
  [CmdletBinding()]param	( [Parameter(Mandatory = $True)] [string]$Source,
    [Parameter(Mandatory = $True)] [string]$Destination )
  mkdir $Destination -Force | Out-Null 
  $FOF_CREATEPROGRESSDLG = "&H0&"  #$FOF_SILENT_FLAG = 4 $FOF_NOCONFIRMATION_FLAG = 16
  $objShell = New-Object -ComObject "Shell.Application"  
  $objFolder = $objShell.NameSpace($Destination).CopyHere($Source, 16)
}

Function MoveWin {
  [CmdletBinding()]param	( [Parameter(Mandatory = $True)] [string]$Source,
    [Parameter(Mandatory = $True)] [string]$Destination )
  mkdir $Destination -Force | Out-Null 
  $objShell = New-Object -ComObject "Shell.Application"
  $objFolder = $objShell.NameSpace($Destination).MoveHere($Source, 16) #16 - overwrite
}

Function Shortcut {
  [CmdletBinding()]param	( [Parameter(Mandatory = $True)] [string]$Where,
    [Parameter(Mandatory = $True)] [string]$Target )
  $s = (New-Object -COM WScript.Shell).CreateShortcut($Where)
  $s.TargetPath = $Target
  $s.Save()
}

function Get-BoundParam {
  ($(foreach ($bp in $Global:MyInvocation.BoundParameters.GetEnumerator()) {
      # argument list
      $valRep =
      if ($bp.Value -is [switch]) {
        # switch parameter
        if ($bp.Value) { $sep = '' } # switch parameter name by itself is enough
        else { $sep = ':'; '$false' } # `-switch:$false` required
      }
      else {
        # Other data types, possibly *arrays* of values.
        $sep = ' '
        foreach ($val in $bp.Value) {
          if ($val -is [bool]) {
            # a Boolean parameter (rare)
              ('$false', '$true')[$val] # Booleans must be represented this way.
          }
          else {
            # all other types: stringify in a culture-invariant manner.
            if (-not ($val.GetType().IsPrimitive -or $val.GetType() -in [string], [datetime], [datetimeoffset], [decimal], [bigint])) {
              Write-Warning "Argument of type [$($val.GetType().FullName)] will likely not round-trip correctly; stringifies to: $val"
            }
            # Single-quote the (stringified) value only if necessary
            # (if it contains argument-mode metacharacters).
            if ($val -match '[ $''"`,;(){}|&<>@#]') { "'{0}'" -f ($val -replace "'", "''") }
            else { "$val" }
          }
        }
      }
      # Synthesize the parameter-value representation.
      '-{0}{1}{2}' -f $bp.Key, $sep, ($valRep -join ', ')
    }) -join ' ') # join all parameter-value representations with spaces
}

function Get-UnboundParam {
  $Global:MyInvocation.UnboundArguments.GetEnumerator() | % { """$_""" }
}

function Admin {
  #[environment]::GetCommandLineArgs()
  Init
  pushd "$ScriptPath"
  if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { 
    Start-Process powershell.exe -Verb RunAs "-NoProfile -ExecutionPolicy Bypass -File `"$(Get-CallingFileName)`" $(Get-BoundParam) $(Get-UnboundParam)" ; exit 
  }
  popd
}

function AdminLL {
  if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) 
  { Start-Process powershell.exe -Verb RunAs "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" ; exit }
}

function AdminLocal {
  # Working - not from module, copy code to your ps1 file
  if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) 
  { Start-Process powershell.exe -Verb RunAs "-NoProfile -ExecutionPolicy Bypass -File `"$(GetUnc $PSCommandPath)`"" ; exit }
}

function Admin2 {
  If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    if ($args.Count -eq 1) { $arguments = '-ExecutionPolicy Bypass -File "' + (GetUnc $args[0]) + '"' }
    else { $arguments = '-ExecutionPolicy Bypass -File "' + (GetUnc $ScriptPath) + '"' }
    Start-Process powershell -Verb runAs -ArgumentList $arguments; Sleep -s 1; Exit
  }
}

function Notify-Baloon {
  Add-Type -AssemblyName System.Windows.Forms 
  $global:balloon = New-Object System.Windows.Forms.NotifyIcon
  $path = (Get-Process -id $pid).Path
  $balloon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path) 
  $balloon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Warning 
  $balloon.BalloonTipText = 'Your computer will be restarted, please save all your work !!!'
  $balloon.BalloonTipTitle = "Attention $Env:USERNAME" 
  $balloon.Visible = $true 
  $balloon.ShowBalloonTip(5000)
}

function Wait4Key {
  Write-Host -NoNewLine "Press any key to continue...";
  $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown");
}

function InputBox {
  $input = $(Add-Type -AssemblyName Microsoft.VisualBasic
    [Microsoft.VisualBasic.Interaction]::InputBox('Provide name or number', 'Prompt', '58691') )
}

function MessageBox {
  [reflection.assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
  [System.Windows.Forms.Application]::EnableVisualStyles()
  [System.Windows.Forms.MessageBox]::Show("Would you like a MessageBox popup ?", "This is a question !", "YesNoCancel") #"Ok” , "Error”, "AbortRetryIgnore” , "Warning”
  [System.Windows.Forms.MessageBox]::Show("Would you like a MessageBox popup ?", "This is a warning !", "AbortRetryIgnore" , "Warning")
  [Enum]::GetNames([System.Windows.Forms.MessageBoxIcon])
  [Enum]::GetNames([System.Windows.Forms.MessageBoxButtons])
}

function Popup {
  $wshell = New-Object -ComObject Wscript.Shell
  $wshell.Popup($args[0], 0, "Done", 0x1)
}

function RemotePopup($pc, $text) {
  Invoke-WmiMethod -Class Win32_Process -ComputerName $pc -Name Create -ArgumentList "C:\Windows\System32\msg.exe * $text"
}

function Get-LoggedUser1 {
  # WMI shows only local logins
  param([Parameter(Mandatory = $True, ValueFromPipeline = $True, ValueFromPipelinebyPropertyName = $true)]
    [alias("CN", "MachineName", "Device Name")]
    [string]$ComputerName	
  )
  @(Get-WmiObject -ComputerName $ComputerName -Namespace root\cimv2 -Class Win32_ComputerSystem)[0].UserName.Split('\')[1]
  #@(Get-WmiObject –ComputerName $ComputerName –Class Win32_ComputerSystem)[0].Username.Split('\')[1]
  #(Get-ChildItem "c:\Users" | Sort-Object LastWriteTime -Descending | Select-Object Name, LastWriteTime -first 2).Name
  #Get-WmiObject -Class Win32_ComputerSystem -Property UserName -ComputerName .
}

function Get-LoggedUser2 {
  #checks explorer owner
  param	(
    #[Parameter(Mandatory=$True,
    #ValueFromPipeline=$True, ValueFromPipelinebyPropertyName=$true)]
    [alias("CN", "MachineName", "Device Name")]
    [string]$ComputerName	
  )
  If ([string]::IsNullOrEmpty($ComputerName)) { [string]$ComputerName = (Read-Host "Enter a hostname or IP ") }  
  $pc = $ComputerName
  $explorerprocesses = @(Get-WmiObject -Query "Select * FROM Win32_Process WHERE Name='explorer.exe'" -ErrorAction SilentlyContinue -ComputerName $pc)
  If ($explorerprocesses.Count -eq 0) {
    "No explorer process found / Nobody interactively logged on"
  }
  Else {
    ForEach ($i in $explorerprocesses) {
      $Username = $i.GetOwner().User
      $Domain = $i.GetOwner().Domain
      Write-Host "$Domain\$Username logged on since: $($i.ConvertToDateTime($i.CreationDate))" 
    }
  }
}

function ShortcutUSB {
  $AppLocation = "C:\Windows\System32\rundll32.exe"
  $WshShell = New-Object -ComObject WScript.Shell
  $Shortcut = $WshShell.CreateShortcut("$Home\Desktop\USB Hardware.lnk")
  $Shortcut.TargetPath = $AppLocation
  $Shortcut.Arguments = "shell32.dll,Control_RunDLL hotplug.dll"
  $Shortcut.IconLocation = "hotplug.dll,0"
  $Shortcut.Description = "Device Removal"
  $Shortcut.WorkingDirectory = "C:\Windows\System32"
  $Shortcut.Save()
}

function Test-ComputerConnection {
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $True,	ValueFromPipeline = $True, ValueFromPipelinebyPropertyName = $true)]
    [alias("CN", "MachineName", "Device Name")]
    [string]$ComputerName	
  )
  Begin {
    [int]$timeout = 20
    [switch]$resolve = $true
    [int]$TTL = 128
    [switch]$DontFragment = $false
    [int]$buffersize = 32
    $options = new-object system.net.networkinformation.pingoptions
    $options.TTL = $TTL
    $options.DontFragment = $DontFragment
    $buffer = ([system.text.encoding]::ASCII).getbytes("a" * $buffersize)	
  }
  Process {
    $ping = new-object system.net.networkinformation.ping
    try { $reply = $ping.Send($ComputerName, $timeout, $buffer, $options) }
    catch { $ErrorMessage = $_.Exception.Message }
    if ($reply.status -eq "Success") {
      $props = @{ComputerName = $ComputerName; Online = $True }
    }
    else	{
      $props = @{ComputerName = $ComputerName; Online = $False }
    }
    New-Object -TypeName PSObject -Property $props
  }
  End { }
}

function Get-IpByName($PCname) {
  [System.Net.Dns]::GetHostByName($PCname).AddressList.IPAddressToString
}

function Get-HostByIP($IP) {
  [System.Net.Dns]::GetHostbyAddress($IP) 
}

function Get-Mac($pc) {
  Get-WmiObject -ClassName Win32_NetworkAdapterConfiguration -Filter "IPEnabled='True'" -ComputerName $pc | 
  Select-Object -Property __SERVER, IPAddress, MACAddress, Description
}

#static $cred = Get-Credential 
function Get-MacAdm($pc) {
  Get-WmiObject -ClassName Win32_NetworkAdapterConfiguration -Filter "IPEnabled='True'" -ComputerName $pc -Credential $cred |
  Select-Object -Property __SERVER, IPAddress, MACAddress, Description
}

function Get-Displays($pc) {
  (Get-WmiObject -ComputerName $pc win32_VideoController).name
  Get-WmiObject -ComputerName $pc WmiMonitorID -Namespace root\wmi | Select @{n = "Connected To"; e = { ($_.__Server) } }, @{n = "Make_Model"; e = { [System.Text.Encoding]::ASCII.GetString($_.UserFriendlyName -ne 00) } }, @{n = "Serial Number"; e = { [System.Text.Encoding]::ASCII.GetString($_.SerialNumberID -ne 00) } } | Out-GridView
}

function Accelerators {
  $TAType = [psobject].Assembly.GetType("System.Management.Automation.TypeAccelerators")
  $TAType::Add('accelerators', $TAType)

  [accelerators]::Get   # this now works
}

function Set-Foreground($hWnd) {
  $pinvokes = @'
  [DllImport("user32.dll", CharSet=CharSet.Auto)]
  public static extern IntPtr FindWindow(IntPtr sClassName, string lpWindowName);
  [DllImport("user32.dll")]
  [return: MarshalAs(UnmanagedType.Bool)]
  public static extern bool SetForegroundWindow(IntPtr hWnd);
'@
  Add-Type -MemberDefinition $pinvokes -Name My -Namespace MB
  # [MB.My]::FindWindow([intptr]::zero,"Administrator: Windows PowerShell")
  [MB.My]::SetForegroundWindow($hWnd)
}

function Split-File {
  $i = 0; Get-Content CBS.log -ReadCount 10000 | % { $i++; $_ | Out-File out_$i.txt }
}

function Trace-Expression {
  # New-Alias -Name tre -Value Trace-Expression -Force #Export-ModuleMember -Function * -Alias *
  [CmdletBinding(DefaultParameterSetName = 'Host')]
  param (
    # ScriptBlock that will be traced.
    [Parameter(
      ValueFromPipeline = $true,
      Mandatory = $true,
      HelpMessage = 'Expression to be traced'
    )]
    [ScriptBlock]$Expression,

    # Name of the Trace Source(s) to be traced.
    [Parameter(
      Mandatory = $true,
      HelpMessage = 'Name of trace, see Get-TraceSource for valid values'
    )]
    [ValidateScript( {
        Get-TraceSource -Name $_ -ErrorAction Stop
      })]
    [string[]]$Name,

    # Option to leave only trace information
    # without actual expression results.
    [switch]$Quiet,

    # Path to file. If specified - trace will be sent to file instead of host.
    [Parameter(ParameterSetName = 'File')]
    [ValidateScript( {
        Test-Path $_ -IsValid
      })]
    [string]$FilePath
  )

  begin {
    if ($FilePath) {
      # assume we want to overwrite trace file
      $PSBoundParameters.Force = $true
    }
    else {
      $PSBoundParameters.PSHost = $true
    }
    if ($Quiet) {
      $Out = Get-Command Out-Null
      $PSBoundParameters.Remove('Quiet') | Out-Null
    }
    else {
      $Out = Get-Command Out-Default
    }
  }

  process {
    Trace-Command @PSBoundParameters | &amp; $Out
  }
}

function Get-InstalledSoftware {
  <#
    .SYNOPSIS
        Retrieves a list of all software installed
    .EXAMPLE
        Get-InstalledSoftware
        
        This example retrieves all software installed on the local computer
    .PARAMETER Name
        The software title you'd like to limit the query to.
    #>
  [OutputType([System.Management.Automation.PSObject])]
  [CmdletBinding()]
  param (
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$Name
  )

  $UninstallKeys = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall", "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
  $null = New-PSDrive -Name HKU -PSProvider Registry -Root Registry::HKEY_USERS
  $UninstallKeys += Get-ChildItem HKU: -ErrorAction SilentlyContinue | Where-Object { $_.Name -match 'S-\d-\d+-(\d+-){1,14}\d+$' } | ForEach-Object { "HKU:\$($_.PSChildName)\Software\Microsoft\Windows\CurrentVersion\Uninstall" }
  if (-not $UninstallKeys) {
    Write-Verbose -Message 'No software registry keys found'
  }
  else {
    foreach ($UninstallKey in $UninstallKeys) {
      if ($PSBoundParameters.ContainsKey('Name')) {
        $WhereBlock = { ($_.PSChildName -match '^{[A-Z0-9]{8}-([A-Z0-9]{4}-){3}[A-Z0-9]{12}}$') -and ($_.GetValue('DisplayName') -like "$Name*") }
      }
      else {
        $WhereBlock = { ($_.PSChildName -match '^{[A-Z0-9]{8}-([A-Z0-9]{4}-){3}[A-Z0-9]{12}}$') -and ($_.GetValue('DisplayName')) }
      }
      $gciParams = @{
        Path        = $UninstallKey
        ErrorAction = 'SilentlyContinue'
      }
      $selectProperties = @(
        @{n = 'GUID'; e = { $_.PSChildName } }, 
        @{n = 'Name'; e = { $_.GetValue('DisplayName') } }
      )
      Get-ChildItem @gciParams | Where $WhereBlock | Select-Object -Property $selectProperties
    }
  }
}

function Get-InstalledApp2 {
  [cmdletbinding()]            
  param(            
    [parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]            
    [string[]]$ComputerName, #(Get-Content list.txt),       #$env:computername,   
    [String[]]$Name
  )            
            
  begin {   
    if (-not $ComputerName) { if (-not (Test-path list.txt)) { $ComputerName = (Get-ADComputer -Filter { OperatingSystem -NotLike "*server*" }).name } else { $ComputerName = Get-Content list.txt } }
    $UninstallRegKeys = @("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall",            
      "SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall")           
  }            
            
  process { 
    $i = 0          
    foreach ($Computer in $ComputerName) {  
      $perc = [math]::Round($i / $ComputerName.Count * 100, 1)
      Write-Progress "Getting Information from computer $Computer" "Complete : $perc %" -perc $perc; $i++   
      Write-Verbose "Working on $Computer"            
      if (Test-Connection -ComputerName $Computer -Count 1 -ea 0) {            
        foreach ($UninstallRegKey in $UninstallRegKeys) {            
          try {            
            $HKLM = [microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine', $computer)            
            $UninstallRef = $HKLM.OpenSubKey($UninstallRegKey)            
            $Applications = $UninstallRef.GetSubKeyNames()            
          }
          catch {            
            Write-Verbose "Failed to read $UninstallRegKey"            
            Continue            
          }            
            
          foreach ($App in $Applications) {     
            foreach ($Nam in $Name) {   
              $AppRegistryKey = $UninstallRegKey + "\\" + $App            
              $AppDetails = $HKLM.OpenSubKey($AppRegistryKey)            
              $AppGUID = $App            
              $AppDisplayName = $($AppDetails.GetValue("DisplayName"))  
              if ($AppDisplayName -notlike $Nam) { continue }
              $AppVersion = $($AppDetails.GetValue("DisplayVersion"))            
              $AppPublisher = $($AppDetails.GetValue("Publisher"))            
              $AppInstalledDate = $($AppDetails.GetValue("InstallDate"))            
              $AppUninstall = $($AppDetails.GetValue("UninstallString"))            
              if ($UninstallRegKey -match "Wow6432Node") {            
                $Softwarearchitecture = "x86" 
              }
              else { $Softwarearchitecture = "x64" }            
              if (!$AppDisplayName) { continue }            
              $OutputObj = New-Object -TypeName PSobject             
              $OutputObj | Add-Member -MemberType NoteProperty -Name ComputerName -Value $Computer.ToUpper()            
              $OutputObj | Add-Member -MemberType NoteProperty -Name AppName -Value $AppDisplayName            
              $OutputObj | Add-Member -MemberType NoteProperty -Name AppVersion -Value $AppVersion            
              $OutputObj | Add-Member -MemberType NoteProperty -Name AppVendor -Value $AppPublisher            
              $OutputObj | Add-Member -MemberType NoteProperty -Name InstalledDate -Value $AppInstalledDate            
              $OutputObj | Add-Member -MemberType NoteProperty -Name UninstallKey -Value $AppUninstall            
              $OutputObj | Add-Member -MemberType NoteProperty -Name AppGUID -Value $AppGUID            
              $OutputObj | Add-Member -MemberType NoteProperty -Name SoftwareArchitecture -Value $Softwarearchitecture            
              $OutputObj     
              $all += , $OutputObj 
            }
          }            
        }             
      }
      else {
        $OutputObj = New-Object -TypeName PSobject             
        $OutputObj | Add-Member -MemberType NoteProperty -Name ComputerName -Value $Computer.ToUpper()            
        $OutputObj | Add-Member -MemberType NoteProperty -Name AppName -Value "OFFLINE" 
        $OutputObj 
        $all += , $OutputObj
      }     
    }            
  }            
            
  end {}
}

function Get-InstalledApp {
  [cmdletbinding()]            
  param(            
    [parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]            
    [string[]]$ComputerName = $env:COMPUTERNAME,   
    [String[]]$Name
  )            
            
  begin {   
    if (-not $ComputerName) {
      #if (-not (Test-path list.txt)) { $ComputerName = (Get-ADComputer -Filter {OperatingSystem -NotLike "*server*"}).name } else { $ComputerName = Get-Content list.txt } 
  
    }
    $UninstallRegKeys = @("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall",            
      "SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall")           
  }            
            
  process { 
    $i = 0          
    foreach ($Computer in $ComputerName) {  
      $perc = [math]::Round($i / $ComputerName.Count * 100, 1)
      Write-Progress "Getting Information from computer $Computer" "Complete : $perc %" -perc $perc; $i++   
      Write-Verbose "Working on $Computer"            
      if (Aping $Computer) {            
        foreach ($UninstallRegKey in $UninstallRegKeys) {            
          try {        
            $HKLM = [microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine', $computer)            
            $UninstallRef = $HKLM.OpenSubKey($UninstallRegKey)            
            $Applications = $UninstallRef.GetSubKeyNames()            
          }
          catch { Write-Verbose "Failed to read $UninstallRegKey"; Continue }            
            
          foreach ($App in $Applications) {     
            foreach ($Nam in $Name) {   
              $AppRegistryKey = $UninstallRegKey + "\\" + $App            
              $AppDetails = $HKLM.OpenSubKey($AppRegistryKey)                       
              $AppDisplayName = $($AppDetails.GetValue("DisplayName"))  
              if (!$AppDisplayName -or $AppDisplayName -notlike $Nam) { continue }                         
              [PSCustomObject]@{
                ComputerName         = $Computer.ToUpper();
                AppName              = $AppDisplayName;
                AppVersion           = $AppDetails.GetValue("DisplayVersion");
                AppVendor            = $AppDetails.GetValue("Publisher");
                InstalledDate        = $AppDetails.GetValue("InstallDate");
                InstallLocation      = $AppDetails.GetValue("InstallLocation");
                InstallSource        = $AppDetails.GetValue("InstallSource");
                URLInfoAbout         = $AppDetails.GetValue("URLInfoAbout");
                UninstallKey         = $AppDetails.GetValue("UninstallString");
                AppGUID              = $AppGUID = $App;
                RegKey               = $AppRegistryKey -replace '\\\\', '\'
                SoftwareArchitecture = if ($UninstallRegKey -match "Wow6432Node") { "x86" } else { "x64" }    
              }
            }
          }
        }             
      }
      else { [PSCustomObject]@{ ComputerName = $Computer.ToUpper(); AppName = 'OFFLINE'; } }     
    }
  }                      
  end {}
}

function Is-Installed {
  [CmdletBinding()]
  Param  ( [Parameter(Mandatory = $True)] [string]$name )
  $x86 = ((Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall") |
    Where-Object { $_.GetValue( "DisplayName" ) -like "*$name*" } ).Length -gt 0;
  $x64 = ((Get-ChildItem "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall") |
    Where-Object { $_.GetValue( "DisplayName" ) -like "*$name*" } ).Length -gt 0;
  return $x86 -or $x64;
}

function Is-InstalledRemoteWMIdonotuse { 
  [CmdletBinding()]
  Param( [Parameter(Mandatory = $True)] [string]$program, [Parameter(Mandatory = $True)] [string]$rhost )
  Try { Get-WMIObject -Class win32_product -Filter "Name like '%$program%'" -ComputerName $rhost -ErrorAction STOP | Select-Object -Property $rhost, Name, Version }
  Catch { Write-Output "$rhost Offline " }
}

Function Show-MeBeingSuperBusy {
  [CmdletBinding()]
  Param (
    [Parameter()]
    [ValidateRange(1, 10)]
    [int]$ConsoleCount = 4
  )
    
  Begin {
    $Argument = '-NoProfile -Command & {1..50 | ForEach-Object {Get-PSDrive}}',
    '-NoProfile -Command & {1..50 | ForEach-Object {Get-Process}}',
    '-NoProfile -Command & {1..50 | ForEach-Object {Get-Service}}',
    '-NoProfile -Command & {1..50 | ForEach-Object {Get-Item -Path env:\}}'
  } # End Begin.
    
  Process {
    For ($i = 1; $i -le $ConsoleCount; $i++) {
      Start-Process -FilePath powershell.exe -ArgumentList ($Argument | Get-Random)
    } # End For.
  } # End Process.
    
  End {
  } # End End.
}

function ShortKeySetup {
  Set-PSReadLineKeyHandler -Key ctrl+B  -BriefDescription 'show busy' -LongDescription "make it look like I am working" -ScriptBlock {
    param($key, $arg)
    #Add-Type -Assembly PresentationCore
    [Microsoft.PowerShell.PSConsoleReadLine]::RevertLine();
    [Microsoft.PowerShell.PSConsoleReadLine]::Insert('Show-MeBeingSuperBusy -ConsoleCount 3; clear;');
    [Microsoft.PowerShell.PSConsoleReadLine]::AcceptLine();
  }
}

function Get-Bios($pc) {
  $opt = New-CimSessionOption -Protocol DCOM
  $s = New-CimSession -Computername $pc -SessionOption $opt -ErrorAction Stop
  Get-CimInstance Win32_bios -CimSession $s
  Remove-CimSession $s
}

function Get-Ram($pc) {
  $opt = New-CimSessionOption -Protocol DCOM
  $s = New-CimSession -Computername $pc -SessionOption $opt -ErrorAction Stop
  $total = 0
  $ram = (Get-CimInstance cim_physicalmemory -CimSession $s | % { [String]($_.Capacity / 1024MB) } )
  $ram | % { $total = $total + $_ } 
  Remove-CimSession $s
  [PSCustomObject]@{ pc = $pc; RAM = $ram -join ','; Total = $total }
}

function Get-Info {
  [CmdletBinding()]
  param (
    [Parameter(ValueFromPipeline = $True, ValueFromPipelinebyPropertyName = $true)] #Mandatory=$True,
    [alias("ComputerName", "MachineName", "DeviceName")]
    [string]$cn = $env:computername )
  SetEmpty; $Err = $false
  $opt = New-CimSessionOption -Protocol DCOM
  if (APing($cn)) {
    # $culture = [Globalization.CultureInfo]::InvariantCulture
    try { $s = New-CimSession -Computername $cn -SessionOption $opt -ErrorAction Stop }   # $ErrorActionPreference = Stop
    catch { $Err = $True }
    if (!$Err) {
      $gcs = Get-CimInstance Win32_ComputerSystem -CimSession $s
      $gos = Get-CimInstance Win32_OperatingSystem -CimSession $s
      $o.ip = (Get-CimInstance Win32_NetworkAdapterConfiguration -CimSession $s).where( { $_.DefaultIPGateway -ne $null }).IPAddress -join ', '
      $o.ramP = (Get-CimInstance cim_physicalmemory -CimSession $s | % { [String]($_.Capacity / 1024MB) } ) -join ','                                     #speed, formfactor, manufacturer
      $o.net = (Get-CimInstance win32_networkadapter -CimSession $s -filter "netconnectionstatus = 2").name -join ', ' -replace ' Virtual Ethernet Adapter'
      $o.hdd = (Get-CimInstance win32_logicaldisk -CimSession $s -Filter "DriveType=3" | select @{l = 'Size'; e = { [math]::Round(($_.Size / 1GB), 1) } }).size -join ', '
      $o.dvd = (Get-CimInstance Win32_CDROMDrive -CimSession $s).Caption -join ', '
      $o.vid = (Get-CimInstance Win32_VideoController -CimSession $s).name -join ', ' -replace " \(Microsoft Corporation .*?\)"
      $o.serial = (Get-CimInstance Win32_bios -CimSession $s).SerialNumber
      #Remove-CimInstance -InputObject $gcs

      $o.UpdTime = (Get-Date).ToString('HH:mm:ss')
      $o.UpdDate = (Get-Date).ToString('dd/MM/yyyy')
      $o.host = $ComputerName
      if ($gcs.Username) { $o.user = $gcs.Username.Split('\')[1] } else { $o.dn = 'NOT LOGGED' }
      try { if ($o.user) { $o.dn = [string]([adsisearcher]"(&(objectCategory=person)(objectClass=user)(samaccountname=$($o.user)))").FindOne().Properties['displayname'] } } 
      catch { $o.dn = "PRD\$($o.user)"; $o.user = $o.dn }
      $o.model = $gcs.Model -replace "OptiPlex " -replace "Precision " -replace "WorkStation " -replace "Tower " -replace " Virtual Platform"
      $o.ram = ($gcs | select @{l = 'RAM'; e = { [math]::Round(($_.TotalPhysicalMemory / 1GB), 0) } }).Ram
      Switch -Wildcard ($gos.caption) {
        "Microsoft Windows 10 *" { $o.os = 'Win 10'; break }
        "Microsoft Windows 7 *" { $o.os = 'Win 7'; break }
      }
      $o.bit = $gos.OSArchitecture  #win32_cs SystemType
      if ($o.serial -like "*VMware*" ) { $o.serial = "VMware" } 
      $regKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $CN)
      $regKey = $regKey.OpenSubKey('SOFTWARE\Microsoft\Windows NT\CurrentVersion')
      if ($o.os -like "*Win 10*") { $o.ver = $regKey.GetValue('ReleaseId') } else { $o.ver = $regKey.GetValue('CurrentBuild') }
      try { $Monitors = @(Get-CimInstance WmiMonitorID -Namespace root\wmi -CimSession $s -ErrorAction Stop) } catch { $Monitors = @() }
      $o.monitors = $Monitors.Count
      $i = 1
      ForEach ($Monitor in $Monitors) {
        $Manufacturer = ($Monitor.ManufacturerName -notmatch '^0$').Foreach( { [char]$_ }) -join ''
        $Name = ($Monitor.UserFriendlyName -notmatch '^0$').Foreach( { [char]$_ }) -join '' -replace "DELL "
        $Serial = ($Monitor.SerialNumberID -notmatch '^0$').Foreach( { [char]$_ }) -join ''
        $exp = "`$o.mon$i = `"$Name`""
        Invoke-Expression $exp
        $i++
      }

    }
    else { $o.dn = "Error" }
    Remove-CimSession $s
  }
  else { $o.dn = "OFFLINE" }
  return $o
}

function Get-MappedDrives($CN) {
  if (APing($cn)) {
    $opt = New-CimSessionOption -Protocol DCOM; $Report = @() 
    try { $s = New-CimSession -Computername $cn -SessionOption $opt -ErrorAction Stop } 
    catch { $Err = $True }
    try { $explorer = Get-CimInstance -CimSession $s -Class win32_process -Filter "name='explorer.exe'" -ErrorAction Stop }
    catch { "$CN - WMI Error " }
    $owner = Invoke-CimMethod -InputObject $explorer -MethodName GetOwner -CimSession $s | Select-Object -ExpandProperty user
    $sid = Invoke-CimMethod -InputObject $explorer -MethodName GetOwnerSid -CimSession $s | Select-Object -ExpandProperty Sid
    if ($explorer) {
      $Hive = [uint32]2147483651                            # [uint32]$hklm = 2147483650   # $khu =  [uint32]2147483651  #wrong [long]$HIVE_HKU = 2147483651 
      $DriveList = Invoke-CimMethod -ClassName 'StdRegProv' -CimSession $s -MethodName 'EnumKey' -Namespace 'ROOT\CIMv2' -Arguments @{hDefKey = $Hive; sSubKeyName = "$($sid)\Network" }
      if ($DriveList.sNames.count -gt 0) {
        #If the SID network has mapped drives iterate and report on said drives
        $Person = $owner
        foreach ($drive in $DriveList.sNames) {
          $hash = [ordered]@{
            ComputerName	= $CN
            User         = $Person
            Drive        = $drive
            Share        = Invoke-CimMethod -ClassName 'StdRegProv' -CimSession $s -MethodName 'GetStringValue' -Namespace 'ROOT\CIMv2' -Arguments @{hDefKey = $Hive; sSubKeyName = "$($sid)\Network\$($drive)"; sValueName = "RemotePath" } | Select-Object -ExpandProperty sValue  # "$(($RegProv.GetStringValue($Hive, "$($sid)\Network\$($drive)", "RemotePath")).sValue)"
          }
          $objDriveInfo = new-object PSObject -Property $hash
          $Report += $objDriveInfo
        }
      }
      else {
        $hash = [ordered]@{
          ComputerName = $CN
          User         = $Person
          Drive        = ""
          Share        = "No mapped drives"
        }
        $objDriveInfo = new-object PSObject -Property $hash
        $Report += $objDriveInfo
      }
    }
    else {
      $hash = [ordered]@{
        ComputerName	= $CN
        User         = "Nobody"
        Drive        = ""
        Share        = "explorer not running"
      }
      $objDriveInfo = new-object PSObject -Property $hash
      $Report += $objDriveInfo
    }
  }
  else {
    $hash = [ordered]@{
      ComputerName	= $CN
      User         = "Nobody"
      Drive        = ""
      Share        = "Cannot connect"
    }
    $objDriveInfo = new-object PSObject -Property $hash
    $Report += $objDriveInfo
  }
  return [array]$Report
}

Function Get-DiskInfo { 
  param($computername = $env:COMPUTERNAME)
  Get-WMIObject Win32_logicaldisk -ComputerName $computername | Select-Object @{Name = 'ComputerName'; Ex = { $computername } }, `
  @{Name = ‘Drive Letter‘; Expression = { $_.DeviceID } }, `
  @{Name = ‘Drive Label’; Expression = { $_.VolumeName } }, `
  @{Name = ‘Size(MB)’; Expression = { [int]($_.Size / 1MB) } }, `
  @{Name = ‘FreeSpace%’; Expression = { [math]::Round($_.FreeSpace / $_.Size, 2) * 100 } }
}  #Get-DiskInfo -computername $WPFtextBox.Text | % {$WPFlistView.AddChild($_)}

function LogonStatus ($computer = 'localhost') {
  $i = 0; $user = $null 
  try { $user = gwmi -Class win32_computersystem -ComputerName $computer | select * -ExpandProperty username -ErrorAction Stop } 
  catch { $i = 1 }                                                                                      #"Not logged on"
  try { if ((Get-Process logonui -ComputerName $computer -ErrorAction Stop) -and ($user)) { $i = 2 } }   #"Workstation locked"
  catch { if ($user) { $i = 3 } }                                                                       #"Computer In Use"
  return $i
} 
 
function APing($PCs) {
  $buffer = ([system.text.encoding]::ASCII).getbytes("a" * [int]32)
  $Task = ForEach ($PC in $PCs) {
    (New-Object System.Net.NetworkInformation.Ping).SendPingAsync($PC, 200, $buffer, @{TTL = 128; DontFragment = $false }) | Add-Member -NotePropertyName Name -NotePropertyValue $pc -PassThru -Force 
  } 
  [void][Threading.Tasks.Task]::WaitAll($Task, 300) 
  $Task.Where( { $_.result.status -eq 'success' }) | % { $_.result | Add-Member -NotePropertyName Name -NotePropertyValue $_.name -Force -ErrorAction SilentlyContinue; $_.result | select * -ExcludeProperty RoundtripTime, Options, Buffer } 
}

function APing2($PCs) {
  $buffer = ([system.text.encoding]::ASCII).getbytes("a" * [int]32)
  $Task = ForEach ($PC in $PCs) {
    (New-Object System.Net.NetworkInformation.Ping).SendPingAsync($PC, 200, $buffer, @{TTL = 128; DontFragment = $false }) | Add-Member -NotePropertyName Name -NotePropertyValue $pc -PassThru -Force 
  } 
  [void][Threading.Tasks.Task]::WaitAll($Task, 200) 
  $Task | % { $_.result | Add-Member -NotePropertyName Name -NotePropertyValue $_.name -Force -ErrorAction SilentlyContinue; $_.result | select * -ExcludeProperty RoundtripTime, Options, Buffer } 
}

Function APingN($PCs) {
  (APing($PCs)).Name
}

function Create-Task {
  $taskname = "Shutdown_task"
  Unregister-ScheduledTask -TaskName $taskname -Confirm:$false -ErrorAction SilentlyContinue
  $RDate = Get-Date -f 'dd/MM/yyyy'                        # 16/03/2016
  $RTime = get-date (get-date).AddMinutes(1) -f 'HH:mm'    # 09:31   +1
  $action = New-ScheduledTaskAction -Execute 'C:\Windows\System32\wscript.exe' -Argument '"C:\Windows\System32\ShutDownTimer.vbs" -interactive' -WorkingDirectory 'C:\Windows\System32\'
  $trigger = New-ScheduledTaskTrigger -Once -at $RTime 
  $Settings = New-ScheduledTaskSettingsSet -Compatibility Win8 
  $user = New-ScheduledTaskPrincipal -GroupId "Users"
  $principal = New-ScheduledTaskPrincipal -UserId (Get-CimInstance –ClassName Win32_ComputerSystem | Select-Object -expand UserName)
  [void](Register-ScheduledTask -Action $action -Trigger $trigger -TaskName $taskname -Description "Shutdown task (interactive)" -Settings $Settings -Principal $principal -Force) 
}

Function old_schTask {   
  $RDate = Get-Date -f 'dd/MM/yyyy'                        # 16/03/2016
  $RTime = get-date (get-date).AddMinutes(1) -f 'HH:mm'    # 09:31   +1
  &schtasks /delete /tn "Shutdown_task" /f 
  &schtasks /create /sc once /RU "USERS" /tn "Shutdown_task" /tr "'C:\Windows\System32\wscript.exe' C:\Windows\System32\ShutDownTimer.vbs -interactive" /SD $RDate /ST $RTime /f /RL HIGHEST /IT
}

function Loge($text) {
  New-EventLog –LogName Application –Source “MBmod Script” -ErrorAction SilentlyContinue 
  Write-EventLog –LogName Application –Source “MBmod Script” –EntryType Information –EventID 1 –Message $text
}

Function pause1 ($message) {
  if ($psISE) {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show("$message")
  }
  else {
    Write-Host "$message" -ForegroundColor Yellow
    $x = $host.ui.RawUI.ReadKey("NoEcho,IncludeKeyDown")
  }
}

#helpers to hide the console window
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
' 
function Show-Console {
  $consolePtr = [Console.Window]::GetConsoleWindow()
  [void][Console.Window]::ShowWindow($consolePtr, 4)
}
function Hide-Console {
  $consolePtr = [Console.Window]::GetConsoleWindow()
  [void][Console.Window]::ShowWindow($consolePtr, 0) #0 hide
}
#/Helpers

function Is-Numeric ($Value) {
  return $Value -match "^[\d\.*]+$" #"^[\d\.]+$"
}

function Is-Alpha ($Value) {
  return $Value -match '^[a-z''*]+$' #'^[a-z\s]+$'
}

function ask-old {
  $ok = $false
  $fnd = @()
  $SearchBase = 'LDAP://DC=dealers,DC=aib,DC=pri'
  $Props = ('displayname', 'samaccountname', 'givenname', 'sn')
  $ads = [adsisearcher]"()"                               ###### Prepare [Active Directory Searcher]
  $ads.searchRoot = [adsi]$global:SearchBase
  $ads.PropertiesToLoad.AddRange($global:props) | Out-Null


  Do { 
    "`n`t`tPlease provide name, surname or staff ID."
    "`t`tYou can specify wildcard characters * "
    $inp = Read-Host -Prompt "`t`t(leave blank for yourself)"
    if (!$inp) { $inp = [Environment]::UserName }

    if (Is-Numeric $inp) { $ads.Filter = "(&(objectCategory=person)(objectClass=user)(samaccountname=$inp))" } 
    elseif (Is-Alpha $inp) { $ads.Filter = "(&(objectCategory=person)(objectCategory=User)(|(sn=$inp)(givenname=$inp)))" }                  #$fnames+=([adsisearcher]"(&(objectCategory=person)(objectCategory=User)(givenname=$inp))").FindAll().Properties.displayname  #$snames+=([adsisearcher]"(&(objectCategory=person)(objectCategory=User)(sn=$inp))").FindAll().Properties.displayname
    else { Write-Host; Write-Warning "Wrong Stuff ID"; continue }  

    $fnd = $ads.FindAll() 
    #$fnd.count
    if ($fnd.count -gt 0) { $ok = $true } else { Write-Host; Write-Warning "No user $inp in Active Directory" }  
  } while (!$ok)

  $list = @()
  foreach ($f in $fnd) {
    if (Is-Numeric $f.properties.samaccountname) {
      $list += [PSCustomObject]@{
        fn   = [string]$f.properties.givenname
        sn   = [string]$f.properties.sn
        dn   = [string]$f.properties.displayname
        sam  = [string]$f.properties.samaccountname
        loc  = [string]''
        host = [string]''
      }
    }
  }
  ''
  ($list | select dn, sam | ft -HideTableHeaders | Out-String).Trim()
  ''
}

function CheckInput {
  $pos = $Host.UI.RawUI.CursorPosition # @{X=$x;Y=$y}
  $u = 0
  do {
    
    $txt = if ($u.Count -lt 2) { "Search " } else { "[1-$($u.count)] or search " }
    $inp = (Read-Host -Prompt $txt).Trim()
    $ok = $False
    #   $Host.UI.RawUI.CursorPosition = $pos
    #   0..$($u.count+5) | %{ $Host.UI.RawUI.CursorPosition = @{X=0;Y=$_} ; $t='     '*30; Write-Host $t }
    #   $Host.UI.RawUI.CursorPosition = $pos
    if (!$inp) { $inp = [Environment]::UserName } #go back line if nothing
    if ($inp -eq 'q' -or $inp -eq 'Q') {
      return [PSCustomObject]@{ L = 'Q' }
      $ok = $true; Break 
    }
    if (Is-Numeric $inp) { 
      if ([int]$inp -le $u.Count) {
        return $u[$inp - 1] 
        $ok = $true; Break 
      } 
    }
    $pos = $Host.UI.RawUI.CursorPosition # @{X=$x;Y=$y}
    [array]$u = SearchAll $inp
    if ($u) {
      Write-host
      ( $u | select L, Name, Desc, Office | ft -HideTableHeaders | Out-String).TrimEnd() | Out-Host
      Write-host   
    }
    else { Write-host "- Nothing found with - $inp" }
    if ($u.count -eq 1) { $ok = $true; return $u }
  } while (!$ok) 
}

function UnifyObj {
  Param([Parameter(ValueFromPipeline)]$O, $L)
  begin { if (-not $l) { $l = 1 } }
  process {
    $IsUser = [bool]($o.PSobject.Properties.name -match "DisplayName")
    [PSCustomObject]@{
      L      = "[$l]"
      Name   = $o.Name
      Desc   = if ($IsUser) { $o.DisplayName } else { $o.Description }    
      Office = if ($IsUser) { $o.Office } else { '' }
      IsPC   = -not $IsUser 
    }
    $l++
  }   
}

function SearchAll {
  param ($inp)
  $in = "*$inp*"; 
  [array]$u = $ADu | ? { $_.Name -like $in -or $_.DisplayName -like $in -or $_.Office -like $in } | select -First 10 | UnifyObj
  $u += $ADc | ? { $_.Name -like $in -or $_.Description -like $in } | select -First 10 | UnifyObj -L ($u.Count + 1)
  return $u
}

function Format-Color([hashtable] $Colors = @{ }, [switch] $SimpleMatch) {
  $lines = ($input | Out-String) -replace "`r", "" -split "`n"
  foreach ($line in $lines) {
    $color = ''
    foreach ($pattern in $Colors.Keys) {
      if (!$SimpleMatch -and $line -match $pattern) { $color = $Colors[$pattern] }
      elseif ($SimpleMatch -and $line -like $pattern) { $color = $Colors[$pattern] }
    }
    if ($color) {
      Write-Host -ForegroundColor $color $line
    }
    else {
      Write-Host $line
    }
  }
}

Function Trace-Word {
  [Cmdletbinding()]
  [Alias("Highlight")]
  Param(
    [Parameter(ValueFromPipeline = $true, Position = 0)] [string[]] $content,
    [Parameter(Position = 1)] 
    [ValidateNotNull()]
    [String[]] $words = $(throw "Provide word[s] to be highlighted!")
  )
  Begin {
    $Color = @{       
      0  = 'Yellow'      
      1  = 'Magenta'     
      2  = 'Red'         
      3  = 'Cyan'        
      4  = 'Green'       
      5  = 'Blue'        
      6  = 'DarkGray'    
      7  = 'Gray'        
      8  = 'DarkYellow'    
      9  = 'DarkMagenta'    
      10 = 'DarkRed'     
      11 = 'DarkCyan'    
      12 = 'DarkGreen'    
      13 = 'DarkBlue'        
    }
    $ColorLookup = @{ }
    For ($i = 0; $i -lt $words.count ; $i++) {
      if ($i -eq 13) { $j = 0 }
      else { $j = $i }
      $ColorLookup.Add($words[$i], $Color[$j])
      $j++
    }
        
  }
  Process {
    $content | ForEach-Object {
      $TotalLength = 0
      $_.split() | `
        # Where-Object {-not [string]::IsNullOrWhiteSpace($_)} | ` #Filter-out whiteSpaces
        ForEach-Object {
        if ($TotalLength -lt ($Host.ui.RawUI.BufferSize.Width - 10)) {
          #"TotalLength : $TotalLength"
          $Token = $_
          $displayed = $False
                            
          Foreach ($Word in $Words) {
            if ($Token -like "*$Word*") {
              $Before, $after = $Token -Split "$Word"
              #"[$Before][$Word][$After]{$Token}`n"
              Write-Host $Before -NoNewline ; 
              Write-Host $Word -NoNewline -Fore Black -Back $ColorLookup[$Word];
              Write-Host $after -NoNewline ; 
              $displayed = $true                                   
              #Start-Sleep -Seconds 1    
              #break  
            }

          } 
          If (-not $displayed) {   
            Write-Host "$Token " -NoNewline                                    
          }
          else {
            Write-Host " " -NoNewline  
          }
          $TotalLength = $TotalLength + $Token.Length + 1
        }
        else {                      
          Write-Host '' #New Line  
          $TotalLength = 0 
        }
        #Start-Sleep -Seconds 0.5
      }
      Write-Host '' #New Line
    }
  }
  end
  { }
}

Function Trace-Word_old {
  [Cmdletbinding()]
  [Alias("Highlight")]
  Param(
    [Parameter(ValueFromPipeline = $true, Position = 0)] [string[]] $content,
    [Parameter(Position = 1)] 
    [ValidateNotNull()]
    [String[]] $words = $(throw "Provide word[s] to be highlighted!")
  )
  Begin {
    $Color = @{       
      0  = 'Yellow'      
      1  = 'Magenta'     
      2  = 'Red'         
      3  = 'Cyan'        
      4  = 'Green'       
      5  = 'Blue'        
      6  = 'DarkGray'    
      7  = 'Gray'        
      8  = 'DarkYellow'    
      9  = 'DarkMagenta'    
      10 = 'DarkRed'     
      11 = 'DarkCyan'    
      12 = 'DarkGreen'    
      13 = 'DarkBlue'        
    }
    $ColorLookup = @{ }
    For ($i = 0; $i -lt $words.count ; $i++) {
      if ($i -eq 13) { $j = 0 }
      else { $j = $i }
      $ColorLookup.Add($words[$i], $Color[$j])
      $j++
    }
        
  }
  Process {
    $content | ForEach-Object {
      $TotalLength = 0
      $_.split() | `
        # Where-Object {-not [string]::IsNullOrWhiteSpace($_)} | ` #Filter-out whiteSpaces
        ForEach-Object {
        if ($TotalLength -lt ($Host.ui.RawUI.BufferSize.Width - 10)) {
          #"TotalLength : $TotalLength"
          $Token = $_
          $displayed = $False
                            
          Foreach ($Word in $Words) {
            if ($Token -like "*$Word*") {
              $Before, $after = $Token -Split "$Word"
              #"[$Before][$Word][$After]{$Token}`n"
              Write-Host $Before -NoNewline ; 
              Write-Host $Word -NoNewline -Fore Black -Back $ColorLookup[$Word];
              Write-Host $after -NoNewline ; 
              $displayed = $true                                   
              #Start-Sleep -Seconds 1    
              #break  
            }

          } 
          If (-not $displayed) {   
            Write-Host "$Token " -NoNewline                                    
          }
          else {
            Write-Host " " -NoNewline  
          }
          $TotalLength = $TotalLength + $Token.Length + 1
        }
        else {                      
          Write-Host '' #New Line  
          $TotalLength = 0 
        }
        #Start-Sleep -Seconds 0.5
      }
      Write-Host '' #New Line
    }
  }
  end
  { }
}

function hostprompt {
  $title = 'Question'
  $question = 'Are you sure you want to proceed?'
  $choices = '&Yes', '&No'

  $decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)
  if ($decision -eq 0) {
    Write-Host 'confirmed'
  }
  else {
    Write-Host 'cancelled'
  }
}

function ToSecureString($plainText) {
  $securestring = new-object System.Security.SecureString
  foreach ($char in ($plainText.toCharArray())) { $secureString.AppendChar($char) }
  return $securestring
}

function Set-Key {
  param([string]$string)
  $length = $string.length
  $pad = 32 - $length
  if (($length -lt 16) -or ($length -gt 32)) { Throw "String must be between 16 and 32 characters" }
  $encoding = New-Object System.Text.ASCIIEncoding
  $bytes = $encoding.GetBytes($string + "0" * $pad)
  return $bytes
}

function Set-EncryptedData {
  param($key, [string]$plainText)
  $securestring = new-object System.Security.SecureString
  foreach ($char in ($plainText.toCharArray())) { $secureString.AppendChar($char) }
  $encryptedData = ConvertFrom-SecureString -SecureString $secureString -Key $key
  return $encryptedData
}

function Get-EncryptedData {
  param($key, $data)
  $data | ConvertTo-SecureString -key $key |
  ForEach-Object { [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($_)) }
}

function EncryptedDataUsage {
  $plainText = "Some Super Secret Password"
  $key = Set-Key "AGoodKeyThatNoOneElseWillKnow"
  $encryptedTextThatIcouldSaveToFile = Set-EncryptedData -key $key -plainText $plaintext
  $encryptedTextThatIcouldSaveToFile
  $DecryptedText = Get-EncryptedData -data $encryptedTextThatIcouldSaveToFile -key $key
  $DecryptedText
}

function RealTimeOutputRedirection {
  param(
    [string] $appFilePath = 'ping.exe',
    [string] $appArguments = 'google.com',
    [string] $appWorkingDirPath = '',
    [int] $consoleOutputEncoding = 0 # 850 = default windows console output encoding (useful for e.g 7z.exe). use "<=0" for host's default encooding.
  )

  if (!$consoleOutputEncoding -le 0) {
    [Console]::OutputEncoding = [System.Text.Encoding]::GetEncoding($consoleOutputEncoding)
  }

  $eventScriptBlock = {
    # received app output
    $receivedAppData = $Event.SourceEventArgs.Data
    # Write output as stream to console in real-time (without -stream parameter output will produce blank lines!)
    #   (without "Out-String" output with multiple lines at once would be displayed as tab delimited line!)
    Write-Host ($receivedAppData | Out-String -Stream)

    <#
        < Insert additional real-time processing steps here.
        < Since it''s in an entirely different scope (not child), variables of parent scope won't be populated to that child scope and scope "$script:" won't work as well. (scope "$global:" would work but should be avoided!)
        < Modify/Enhance variables "*MessageData" (see below) before registering the event to access such variables.
    #>

    # add received data to stringbuilder definded in $stdOutEventMessageData and $stdErrEventMessageData
    $Event.MessageData.outStringBuilder.AppendLine($receivedAppData)
  }

  # MessageData parameters for default events (used for event input and output)
  $stdOutEventMessageData = @{
    # used for adding output within events to stringbuilder (OUT) for further usage
    outStringBuilder = [System.Text.StringBuilder]::new()

    #< add additional properties if necessary. Can be used as input and output in $eventScriptBlock ($Event.MessageData.*)
    #< ....
  }

  # MessageData parameters for error events (used for event input and output)
  $stdErrEventMessageData = @{
    # used for adding output within events to stringbuilder (OUT) for further usage
    outStringBuilder = [System.Text.StringBuilder]::new()

    #< add additional properties if necessary. Can be used as input and output in $eventScriptBlock ($Event.MessageData.*)
    #< ....
  }

  #######################################################
  #region Process-Definition, -Start and Event-Subscriptions (Adaptions in that region should be avoided!)
  #------------------------------------------------------
  try {

    $appProcess = [System.Diagnostics.Process]::new()
    $appProcess.StartInfo = @{
      Arguments              = $appArguments
      WorkingDirectory       = $appWorkingDirPath
      FileName               = $appFilePath # mandatory
      RedirectStandardOutput = $true # mandatory = $true
      RedirectStandardError  = $true # mandatory = $true
      #< RedirectStandardInput  = $true # leave commented; only useful in some circumstances. Didn''t find any use, but mentioned in: https://stackoverflow.com/questions/8808663/get-live-output-from-process
      UseShellExecute        = $false # mandatory = $false
      CreateNoWindow         = $true # mandatory = $true
    }

    $stdOutEvent = Register-ObjectEvent -InputObject $appProcess -Action $eventScriptBlock -EventName 'OutputDataReceived' -MessageData $stdOutEventMessageData
    $stdErrEvent = Register-ObjectEvent -InputObject $appProcess -Action $eventScriptBlock -EventName 'ErrorDataReceived' -MessageData $stdErrEventMessageData

    [void]$appProcess.Start()
    $appProcess.BeginOutputReadLine()
    $appProcess.BeginErrorReadLine()

    while (!$appProcess.HasExited) {
      # Don't use method "WaitForExit()"! This will not show the output in real-time as it blocks the output stream!
      #   using "Sleep" from System.Threading.Thread class for short sleep times below 1/1.5 seconds is better than "Start-Sleep" in terms of PS overhead/performance on init (Test it yourself)
      #   (sleep will block console output --> don't set too high; but also not too low for performance reasons in long running actions)
      [System.Threading.Thread]::Sleep(250)

      #< maybe timeout ...
    }

  }
  finally {
    if (!$appProcess.HasExited) {
      $appProcess.Kill() # WARNING: Entire process gets killed! Review and adapt!
    }

    if ($stdOutEvent -is [System.Management.Automation.PSEventJob]) {
      Unregister-Event -SourceIdentifier $stdOutEvent.Name
    }
    if ($stdErrEvent -is [System.Management.Automation.PSEventJob]) {
      Unregister-Event -SourceIdentifier $stdErrEvent.Name
    }
  }
  #------------------------------------------------------
  #endregion
  #######################################################

  $stdOutText = $stdOutEventMessageData.outStringBuilder.ToString() # final output for further usage
  $stErrText = $stdErrEventMessageData.outStringBuilder.ToString()  # final errors for furt
    
}

function Uninstall-Wmi {
  [cmdletbinding()]            
  param (            
    [parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
    [string]$ComputerName = $env:computername,
    [parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
    [string]$AppGUID
  )            

  try {
    $returnval = ([WMICLASS]"\\$computerName\ROOT\CIMV2:win32_process").Create("msiexec `/x$AppGUID `/norestart `/qn")
  }
  catch {
    write-error "Failed to trigger the uninstallation. Review the error message"
    $_
  }
  switch ($($returnval.returnvalue)) {
    0 { "Uninstallation command triggered successfully" }
    2 { "You don't have sufficient permissions to trigger the command on $Computer" }
    3 { "You don't have sufficient permissions to trigger the command on $Computer" }
    8 { "An unknown error has occurred" }
    9 { "Path Not Found" }
    9 { "Invalid Parameter" }
  }
}

function uninst-java {
  $list = (Get-ADComputer -Filter { OperatingSystem -NotLike "*server*" }).name #(Get-ADComputer -Filter {Name -like "*-bcs"} -SearchBase "OU=CTS Win 10 PC``s,OU=DRS Win 10 PCs,DC=dealers,DC=aib,DC=pri").name
  $on = (aping($list)).name
  rv ii, all -ErrorAction SilentlyContinue
  [System.Collections.ArrayList]$all = @()
  $all = Get-InstalledApp $c "*java 8*"
  Export-Xlsx -obj $all -path 'C:\Users\dsk_58691\Desktop\uninst-java-all.xlsx'
}

function UpdateGraphicDrivers($pc, $drvPath) {

  # Import-Module "H:\MB\PS\modules\MBMod\0.1\MBMod.psm1" -Force -WarningAction SilentlyContinue
  # check for NVIDIA drivers
  # $pc | % { Logged-User $_ } | ft
  # $pc | % { Get-GraphicDrivers $_ | ? { $_.Description -like "*NVIDIA*"} } | sort ComputerName -Unique | sort DriverDate


  $srcfile = split-path $drvPath -Leaf
  $c = "C:\Temp\inst\" + $srcfile + " -s -n Display.Driver"
  $x = 0; $out = @()

  $pc | % { 
    $destPath = "\\$_\c$\Temp\inst\"
    if (-not (test-path "$destPath") ) { md $destPath -Verbose }
    if (-not (test-path "$destPath\$srcfile") ) { Copy-Item $srcPath $destPath -Force -Verbose }
    [PSCustomObject]@{ PC = $_ ; PID = (Run-Remote $_ $c) }
  }

}

function Scan-Updates {
  #Using WUA to Scan for Updates Offline with PowerShell  #VBS version: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/aa387290(v=vs.85)  

  $path = if ($psise) { Split-Path $psise.CurrentFile.FullPath } else { $PSScriptRoot }

  if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) 
  { Start-Process powershell.exe -Verb RunAs "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" ; exit }

  if (Test-Path "$path\wsusscn2.cab") { "File $path\wsusscn2.cab exist" } else {
    "Downloading $path\wsusscn2.cab exist"
    # Turn on proxy for internet access
    Set-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings' -name ProxyServer -Value 'webcorp.prd.aib.pri:8082'
    set-itemproperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings' -name ProxyEnable -value 1 
    #Invoke-WebRequest -Uri "http://download.windowsupdate.com/microsoftupdate/v6/wsusscan/wsusscn2.cab" -OutFile "$path\wsusscn2.cab"
    Start-BitsTransfer -Source "http://download.windowsupdate.com/microsoftupdate/v6/wsusscan/wsusscn2.cab" -Destination "$path\wsusscn2.cab"
    set-itemproperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings' -name ProxyEnable -value 0
  } 
  
  Write-Output "Adding '$path\wsusscn2.cab' to UpdateServiceManager..." 
  $UpdateSession = New-Object -ComObject Microsoft.Update.Session  
  $UpdateServiceManager = New-Object -ComObject Microsoft.Update.ServiceManager  
  $UpdateService = $UpdateServiceManager.AddScanPackageService("Offline Sync Service", "$path\wsusscn2.cab", 1)  
  $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()   
  Write-Output "Searching for updates..."  
  $UpdateSearcher.ServerSelection = 3 #ssOthers 
  $UpdateSearcher.IncludePotentiallySupersededUpdates = $true # good for older OSes, to include Security-Only or superseded updates in the result list, otherwise these are pruned out and not returned as part of the final result list 
  $UpdateSearcher.ServiceID = $UpdateService.ServiceID.ToString()  
  $SearchResult = $UpdateSearcher.Search("IsInstalled=0") # or "IsInstalled=0 or IsInstalled=1" to also list the installed updates as MBSA did  
  $Updates = $SearchResult.Updates  

  $date = (Get-Date -F "yy-MM-dd HH-mm")

  $all = @( $Updates | % { $kb = ($_.Title | Select-String '(?<=\()[^]]+(?=\))' -AllMatches).Matches.Value; [PSCustomObject]@{ KB = $kb; Title = $_.Title } } ) | sort kb -Descending
  $out = $all | % { $_.kb + "`t" + $_.Title }

  if ($Updates.Count -eq 0) {
    "There are no applicable updates." | tee "$path\wsusscan $date.txt"
  }
  else { Write-Output "List of applicable items on the machine when using wssuscan.cab:" }
  
  $out | tee "$path\wsusscan $date.txt" -Append

  function Speak($text) {
    Add-Type -AssemblyName System.speech
    $speak = New-Object System.Speech.Synthesis.SpeechSynthesizer
    $speak.Rate = 3
    $speak.Speak($text) 
  }

  if ( (Get-CimInstance -ClassName Win32_OperatingSystem).ProductType -eq 1) { Speak "Scan Complete" }  #Speak if we are on workstation

  #pause

}

function Get-CMCollectionOfDevice {
  [CmdletBinding()]
  [OutputType([int])]
  Param
  (
    # Computername
    [Parameter(Mandatory = $true,
      ValueFromPipelineByPropertyName = $true,
      Position = 0)]
    [String]$Computer,
 
    # ConfigMgr SiteCode
    [Parameter(Mandatory = $false,
      ValueFromPipelineByPropertyName = $true,
      Position = 1)]
    [String]$SiteCode = "DUB",
 
    # ConfigMgr SiteServer
    [Parameter(Mandatory = $false,
      ValueFromPipelineByPropertyName = $true,
      Position = 2)]
    [String]$SiteServer = "drscmsrv2.dealers.aib.pri"
  )
  Begin {
    [string] $Namespace = "root\SMS\site_$SiteCode"
  }
 
  Process {
    $si = 1
    Write-Progress -Activity "Retrieving ResourceID for computer $computer" -Status "Retrieving data" 
    $ResIDQuery = Get-WmiObject -ComputerName $SiteServer -Namespace $Namespace -Class "SMS_R_SYSTEM" -Filter "Name='$Computer'"
    
    If ([string]::IsNullOrEmpty($ResIDQuery)) {
      Write-Output "System $Computer does not exist in Site $SiteCode"
    }
    Else {
      $Collections = (Get-WmiObject -ComputerName $SiteServer -Class sms_fullcollectionmembership -Namespace $Namespace -Filter "ResourceID = '$($ResIDQuery.ResourceId)'")
      $colcount = $Collections.Count
    
      $devicecollections = @()
      ForEach ($res in $collections) {
        $colid = $res.CollectionID
        Write-Progress -Activity "Processing  $si / $colcount" -Status "Retrieving Collection data" -PercentComplete (($si / $colcount) * 100)
 
        $collectioninfo = Get-WmiObject -ComputerName $SiteServer -Namespace $Namespace -Class "SMS_Collection" -Filter "CollectionID='$colid'"
        $object = New-Object -TypeName PSObject
        $object | Add-Member -MemberType NoteProperty -Name "CollectionID" -Value $collectioninfo.CollectionID
        $object | Add-Member -MemberType NoteProperty -Name "Name" -Value $collectioninfo.Name
        $object | Add-Member -MemberType NoteProperty -Name "Commnent" -Value $collectioninfo.Comment
        $object | Add-Member -MemberType NoteProperty -Name "LastRefreshTime" -Value ([Management.ManagementDateTimeConverter]::ToDateTime($collectioninfo.LastRefreshTime))
        $devicecollections += $object
        $si++
      }
    } # end check system exists
  }
 
  End {
    $devicecollections
  }
}

function Speak($text) {
  Add-Type -AssemblyName System.speech
  $speak = New-Object System.Speech.Synthesis.SpeechSynthesizer
  $speak.Rate = 3
  $speak.Speak($text) 
}

function Get-PCgroup($pc) {
  (Get-ADPrincipalGroupMembership (Get-ADComputer $pc).DistinguishedName).name 
}

function Get-UserGroup($user) {
(Get-ADPrincipalGroupMembership (Get-ADUser $user).DistinguishedName).name
}

function SCCM-ForceUpd($pc) {
  $strAction = "{00000000-0000-0000-0000-000000000121}" # Application Deployment Evaluation Cycle
  try {
    $WMIPath = "\\" + $pc + "\root\ccm:SMS_Client" 
    $SMSwmi = [wmiclass] $WMIPath 
    [Void]$SMSwmi.TriggerSchedule($strAction)
  }
  catch
  { $_.Exception.Message }  
}

function SCCM-refresh($pc) {
  ([wmiclass]"\\$pc\root\ccm:SMS_Client").TriggerSchedule("{00000000-0000-0000-0000-000000000001}")
  Invoke-WMIMethod -ComputerName $pc -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule “{00000000-0000-0000-0000-000000000002}”
  Invoke-WMIMethod -ComputerName $pc -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule “{00000000-0000-0000-0000-000000000003}”
  Invoke-WMIMethod -ComputerName $pc -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule “{00000000-0000-0000-0000-000000000021}”
  # Invoke-WMIMethod -ComputerName $pc -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule “{00000000-0000-0000-0000-000000000102}”
  Invoke-CimMethod -Namespace 'root\ccm' -ClassName 'sms_client' -MethodName TriggerSchedule -Arguments @{sScheduleID = "{00000000-0000-0000-0000-000000000002}" }
}

function Find7050IntelDriver {
  $list = (Get-ADComputer -Filter { OperatingSystem -NotLike "*server*" }).name
  foreach ($pc in $list) {
    $opt = New-CimSessionOption -Protocol DCOM
    try {
      $s = New-CimSession -Computername $pc -SessionOption $opt -ErrorAction Stop -OperationTimeoutSec 2
      $model = (Get-CimInstance Win32_ComputerSystem -CimSession $s -Property Model).model
      if ($model -like "*7050") {
        Get-CimInstance Win32_PnPSignedDriver -Filter 'DeviceName LIKE "Intel(R) Chipset SATA%"' -CimSession $s | % { [PSCustomObject]@{ CN = $pc; DriverVer = $_.DriverVersion; DriverDate = $_.DriverDate; DeviceName = $_.devicename; } }
      }
      Remove-CimSession $s
    }
    catch { } 
  }
  Export-Excel -Path "$env:USERPROFILE\Desktop\RAPID.xlsx" -InputObject $all
}

function Replace-Links($pc, $chromelnk) {
  $path1 = "\\$pc\c$\ProgramData\Microsoft\Windows\Start Menu\Programs\Google Chrome.lnk"
  if (compare (gc $chromelnk) (gc $path1)) { Copy-Item -Path $chromelnk -Destination (Split-Path $path1) -Force -Verbose } else { "Correct - $path1" }
  $userlist = (Get-ChildItem "\\$pc\c$\Users\" -Directory -Exclude Administrator, drwin, public, default*).Name 
  $userlist | % { 
    $p = "\\$pc\c$\Users\$_\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Google Chrome.lnk"  
    if (Test-Path $p) { if (compare (gc $chromelnk) (gc $p)) { Copy-Item -Path $chromelnk -Destination (Split-Path $p) -Force -Verbose } else { "Correct - $p" } }
  }
}

function WOL-IP {
  $Mac = "D8:9E:F3:13:5C:7B"
  $MacByteArray = $Mac -split "[:-]" | ForEach-Object { [Byte] "0x$_" }
  [Byte[]] $MagicPacket = (, 0xFF * 6) + ($MacByteArray * 16)
  $UdpClient = New-Object System.Net.Sockets.UdpClient
  $UdpClient.Connect(([System.Net.IPAddress]::Parse('10.28.222.14')), 7)
  $UdpClient.Send($MagicPacket, $MagicPacket.Length)
  $UdpClient.Close()
}

function Get-Wsus($ServerName = 'drsopsmgr3') {
  [reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration") | out-null
  [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer($ServerName, $false, 8530) 
}

Function GetUpdateState {
  param([string[]]$kbnumber = 'KB5041580', [string]$wsusserver = 'drsopsmgr3', [string]$port = 8530
  )
  $report = @()
  [void][reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration")
  $wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer($wsusserver, $False, $port)
  $CompSc = new-object Microsoft.UpdateServices.Administration.ComputerTargetScope
  $updateScope = new-object Microsoft.UpdateServices.Administration.UpdateScope; 
  $updateScope.UpdateApprovalActions = [Microsoft.UpdateServices.Administration.UpdateApprovalActions]::Install
  foreach ($kb in $kbnumber) {
    #Loop against each KB number passed to the GetUpdateState function 
    $updates = $wsus.GetUpdates($updateScope) | ? { $_.Title -match $kb } #Getting every update where the title matches the $kbnumber
    foreach ($update in $updates) {
      #Loop against the list of updates I stored in $updates in the previous step
      $update.GetUpdateInstallationInfoPerComputerTarget($CompSc) | ? { $_.UpdateApprovalAction -eq "Install" } | % { #for the current update
        #Getting the list of computer object IDs where this update is supposed to be installed ($_.UpdateApprovalAction -eq "Install")
        $Comp = $wsus.GetComputerTarget($_.ComputerTargetId)# using #Computer object ID to retrieve the computer object properties (Name, #IP address)
        $info = "" | select UpdateTitle, LegacyName, SecurityBulletins, Computername, OS , IpAddress, UpdateInstallationStatus, UpdateApprovalAction #Creating a custom PowerShell object to store the information
        $info.UpdateTitle = $update.Title
        $info.LegacyName = $update.LegacyName
        $info.SecurityBulletins = ($update.SecurityBulletins -join ';')
        $info.Computername = $Comp.FullDomainName
        $info.OS = $Comp.OSDescription
        $info.IpAddress = $Comp.IPAddress
        $info.UpdateInstallationStatus = $_.UpdateInstallationState
        $info.UpdateApprovalAction = $_.UpdateApprovalAction
        $report += $info # Storing the information into the $report variable 
      }
    }
  }
  $report | ? { $_.UpdateInstallationStatus -ne 'NotApplicable' -and $_.UpdateInstallationStatus -ne 'Unknown' -and $_.UpdateInstallationStatus -ne 'Installed' } #|  Export-Csv -Path c:\temp\rep_wsus.csv -Append -NoTypeInformation #Filtering the report to list only computers where the updates are not installed
} # Usage: GetUpdateState -kbnumber KB5016616 -wsusserver drsopsmgr2 -port 8530

function Get-PcInfo {
  [cmdletbinding()]
  param(
    [parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
    [string[]]$ComputerName = (Read-Host -Prompt 'Please enter computer name') )
  $ErrorActionPreference = 'silentlycontinue'

  $Apps = "Adobe Acrobat Reader DC", "Citrix online plug-in", "Symantec_EnterpriseVault", "PhishMe Reporter", "Google Chrome",
  "Java 8 Update", "Skype for Business 2016", "Microsoft Office Standard 2013", "QlikView Plugin", "WinZip_", "", "McAfee Agent", "McAfee Endpoint", "Tanium"

  $hostn = $ComputerName                
  $user = $env:username                 #(Get-WmiObject -Class Win32_ComputerSystem | Select-Object UserName).Username.Split('\')[1]
  $file = "H:\Builds\ToDo\${hostn}.txt"

  function showsave($text) {
    $text
    $text >> $file
  }

  $name = (Get-WmiObject Win32_OperatingSystem -ComputerName $ComputerName).caption      #Microsoft Windows 7\10 Enterprise
  $bit = (Get-WmiObject Win32_OperatingSystem -ComputerName $ComputerName).OSArchitecture
  $ver = 0; #(Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ReleaseId
  $build = (gwmi Win32_OperatingSystem -ComputerName $ComputerName).Version 
  if ($build -eq '10.0.18362') { $ver = '19H1' } 
  if ($build -eq '10.0.18363') { $ver = '19H2' } 
  if ($build -eq '10.0.19041') { $ver = '20H1' }
  if ($build -eq '10.0.19042') { $ver = '20H2' } 
  if ($build -eq '10.0.19043') { $ver = '21H1' } 
  if ($build -eq '10.0.19044') { $ver = '21H2' }

  $czas = (Get-Date).ToString('MM/dd/yyyy hh:mm:ss tt')
  $dn = ([adsisearcher]"(&(objectClass=user)(samaccountname=$user))").FindOne().Properties['displayname']
  $ip = (Test-Connection -ComputerName $computername -count 1).IPV4Address.ipaddressTOstring
  $vid = @(Get-WmiObject Win32_VideoController -ComputerName $ComputerName) | ? { $_.name -ne 'DameWare Development Mirror Driver 64-bit' -and $_.name -ne 'Microsoft Remote Display Adapter' }
  if (@($vid).count -gt 1) { $vid = "$($vid[0].name) + $($vid[1].name)" } else { $vid = $vid.name }

  $str = "`r`n -----===== $czas =====----- `r`n`r`n"
  $str += "User      : $user  -  $dn `r`n"
  $str += "Hostname  : $hostn `r`n"
  $str += "IPv4      : $ip `r`n"
  $str += "Serial    : $((Get-WmiObject Win32_bios -ComputerName $ComputerName).SerialNumber) `r`n"
  $str += "Windows   : $name, $bit, $ver `r`n"
  $str += "Model     : $((Get-WmiObject Win32_ComputerSystem -ComputerName $ComputerName).Model) `r`n"
  $str += "BIOS      : $((Get-WmiObject win32_bios).Name) `r`n"
  $str += "Video     : $vid `r`n"
  $str += "RAM       : $((Get-WmiObject Win32_ComputerSystem -ComputerName $ComputerName | select @{l='RAM'; e={[math]::Round(($_.TotalPhysicalMemory / 1GB), 0)}}).Ram) GB `r`n"
  $str += "Network   : $((Get-Wmiobject win32_networkadapter -ComputerName $ComputerName -filter "netconnectionstatus = 2").name) `r`n"
  $str += "HDD       : $((Get-Wmiobject win32_logicaldisk -ComputerName $ComputerName -Filter "DriveType=3" | select @{l='Size'; e={[math]::Round(($_.Size / 1GB), 1)}}).size) GB `r`n"
  $str += "CD/DVD    : $((Get-WmiObject Win32_CDROMDrive -ComputerName $ComputerName).Caption) `r`n"
  showsave($str)

  if (($build -split '\.')[0] -lt 10) { 
    $Monitors = @(Get-WmiObject win32_desktopmonitor);  
    showsave("MonitorNo : $($Monitors.count)`n") 
  } 

  #$tmp = $(Get-PSDrive -PSProvider FileSystem | Where-Object {$_.DisplayRoot -ne $null} | select Name,DisplayRoot | ft -hidetableheaders)
  #$tmp.Count
  #showsave($tmp)

  function numInstances([string]$process) {
    @(Get-Process $process -ErrorAction 0).Count
  }
  <#
$Array = @()
Foreach ($app in $Apps) {
 $Result=[PSCustomObject]@{ Name = $app; IsIns = if ($app) {if ( (Get-InstalledApp $ComputerName "*$app*" | ? { $_.appName -ne 'OFFLINE' } | measure).count -ne 0 ) {$true} else {$false} } }
 $Array += $Result
}
showsave(($Array | Format-Table -HideTableHeaders | Out-String).Trim())
showsave("Tanium process no `t`t: " + $(numInstances("TaniumClient")))
#>
}

function WordFill {

  $template = 'G:\Inventory\DRS Desktop Build & Decommission signoffs\Windows 10 Build Checklist Template.docx'
  $wf = 'C:\Temp\alloc\Windows 10 Build Sheet.docx'
  $fold = 'H:\Builds\ToDo'
  $done = 'H:\Builds\DoneByMe'
  $file = gci $fold *.txt | select -First 1 
  $fn = $file.FullName
  $fn

  function RemoveColon ($fn, $nr) {
    $line = (Get-Content $fn)[$nr]
    $start = $line.IndexOf(':') + 1
    $result = $line.Substring($start, $line.Length - $start).Trim()
    return $result
  }

  if ( !(Test-Path (Split-Path $wf)) ) { mkdir (Split-Path $wf) | Out-Null }
  if ( !(Test-Path ($wf)) ) { copy $template $wf }

  $l = Get-Content $fn -TotalCount 2  # (Get-Content $fn)[2]
  $time = $l.Replace('-', '').Replace('=', '').Trim()

  $wd = New-Object -ComObject Word.Application 
  $wd.Visible = $fasle
  $Doc = $Wd.Documents.Open($wf)
  #$Doc = $wd.Documents.Open($wordf, $false, $true)
  #$Sel = $Wd.Selection # $sel.StartOf(15)  $sel.MoveDown()

  $t1 = $wd.ActiveDocument.Tables.item(1)
  $t1.Cell(2, 1).Range.Text = RemoveColon $fn 4
  $t1.Cell(2, 2).Range.Text = RemoveColon $fn 8
  $t1.Cell(2, 3).Range.Text = RemoveColon $fn 10
  $t1.Cell(2, 4).Range.Text = RemoveColon $fn 11
  $t1.Cell(2, 5).Range.Text = RemoveColon $fn 12
  $t1.Cell(2, 6).Range.Text = (RemoveColon $fn 13) + "`n" + (RemoveColon $fn 14)
  # $t1.Cell(4,1).Range.Text="Old Hostname"
 
  $t2 = $wd.ActiveDocument.Tables.item(2)
  $t2.Cell(2, 2).Range.Text = (RemoveColon $fn 3).Split('-').Trim()[0] #(RemoveColon $fn 3).substring(0,5)
  $t2.Cell(2, 1).Range.Text = (RemoveColon $fn 3).Split('-').Trim()[1] 
 
  $t3 = $wd.ActiveDocument.Tables.item(3)
  for ($i = 0; $i -lt 11; $i++) { 
    $t3.Cell(3 + $i, 2).Range.Text = (Get-Content $fn)[16 + $i] -split " " | ? { $_ } | select -Last 1 #next tanium and mcaffee
  }
  $t3.Cell(2 + $i, 2).Range.Text = (Get-Content $fn)[27] -split " " | ? { $_ } | select -Last 1 
  $t3.Cell(3 + $i, 2).Range.Text = (Get-Content $fn)[29] -split " " | ? { $_ } | select -Last 1 

  for ($i = 0; $i -lt 6; $i++) { 
    $t3.Cell(38 + $i, 2).Range.Text = "Done"
  }

  $t4 = $wd.ActiveDocument.Tables.item(4)
  $t4.Cell(2, 3).Range.Text = (Get-Content $fn)[28] -split " " | ? { $_ } | select -Last 1 
  for ($i = 0; $i -lt 8; $i++) { 
    $t4.Cell(3 + $i, 3).Range.Text = "Done"
  }

  $t5 = $wd.ActiveDocument.Tables.item(5)
  $t5.Cell(1, 2).Range.text = "Maciej Bonczyk"
  $t5.Cell(1, 3).Range.text = (get-date).ToString("dd/MM/yyyy")

  $saveas = Join-Path $fold -ChildPath ((RemoveColon $fn 4) + '.docx')
  $Doc.SaveAs([REF][system.object]$saveas)

  $wd.quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wd) | Out-Null
  Remove-Variable wd
  [gc]::Collect()
  [gc]::WaitForPendingFinalizers()

  Move-Item $fn (Join-Path $done $file) -Force

  <#
$doc.content.find.execute("Title") #ok
$doc.SaveAs([ref]"c:\work\osreport.docx")
 $sel.font.bold = 0
 $sel.Style ="Title"
 $sel.font.size = 10
 $Sel.ParagraphFormat.Alignment = 1
 $sel.typeText("Nice Title Something else 1 2 3")
 $rng = $doc.Range()
 $rng.Find.Execute("Title")
 $sel.MoveRight()

$selection.EndOf(15)
$selection.MoveDown()
$Word.Selection.TypeText("This text does not belong here")

$Selection.EndKey($END_OF_STORY)
$selection.MoveDown()
$UserTable.AutoFormat(23)
$UserTable.Columns.AutoFit()
$Selection.TypeParagraph()

$Selection.Style = 'Title'
$Selection.TypeText("Hello")
$Selection.TypeParagraph()
$Selection.Style = 'Heading 1'
$Selection.TypeText("Report compiled at $(Get-Date).")
$Selection.TypeParagraph()

$Selection.Font.Bold = 1
$Selection.TypeText('This is Bold')
$Selection.Font.Bold = 0
$Selection.TypeParagraph()
$Selection.Font.Italic = 1
$Selection.TypeText('This is Italic')

$Report = 'C:\Users\proxb\Desktop\ADocument.doc'
$Document.SaveAs([ref]$Report,[ref]$SaveFormat::wdFormatDocument)
$word.Quit()

[Enum]::GetNames([Microsoft.Office.Interop.Word.WdColor]) | ForEach {
    $Selection.Font.Color = $_
    $Selection.TypeText("This is $($_)")
    $Selection.TypeParagraph()    
} 
$Selection.Font.Color = 'wdColorBlack'
$Selection.TypeText('This is back to normal')

[Enum]::GetNames([microsoft.office.interop.word.WdSaveFormat])

[Enum]::GetNames([Microsoft.Office.Interop.Word.WdColor]) | ForEach {
    [pscustomobject]@{Color=$_}
} | Format-Wide -Property Color -Column 4

[Enum]::GetNames([Microsoft.Office.Interop.Word.WdBuiltinStyle]) | ForEach {
    [pscustomobject]@{Style=$_}
} | Format-Wide -Property Style -Column 4

$null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$word)
[gc]::Collect()
[gc]::WaitForPendingFinalizers()
Remove-Variable word 

#>

}

function Backup-UserProfile {
  $PC = Read-Host -Prompt 'Input your computer name'
  $pc.Trim()
(Get-WmiObject -ComputerName $pc -Class Win32_UserProfile).LocalPath | % { $_.split('\')[-1] } | ? { $_ -match "\d" }
  $user = Read-Host -Prompt 'Input the user name'

  $paths = 
  "\\$pc\C$\Users\$($user)\Desktop\",
  "\\$pc\C$\Users\$($user)\Documents\",
  "\\$pc\C$\Users\$($user)\Favorites\",
  "\\$pc\C$\Users\$($user)\AppData\Roaming\Microsoft\Signatures\",
  "\\$pc\C$\Users\$($user)\AppData\Local\Microsoft\Edge\User Data\Default\Bookmarks",
  "\\$pc\C$\Users\$($user)\AppData\Local\Microsoft\Edge\User Data\Default\Collections\",
  "\\$pc\C$\Users\$($user)\AppData\Roaming\Microsoft\Windows\Recent\AutomaticDestinations\"

  $CurrPath = if ($psise) { Split-Path $psise.CurrentFile.FullPath } else { $PSScriptRoot }
  if (Test-Path $paths[0]) {
    $paths | % { if (Test-Path $_) { $_ ; xcopy $_ "$CurrPath\Backups\$($user)\$(Split-Path -Leaf $_)\" /s /f /q /y } }
  }


}

function DisableIPv6Dealers {
  Adinfo
  $list = Ping-DealersPCs
  $list | % { $val = (Get-RemoteReg $_ LocalMachine 'SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters\' 'DisabledComponents')
    if ($val -ne 255) { (Set-RemoteReg $_ LocalMachine 'SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters\' 'DisabledComponents' 255) }
    [PSCustomObject]@{PC = $_; Reg = (Get-RemoteReg $_ LocalMachine 'SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters\' 'DisabledComponents') } }
}

function SCCM-AllActions {
  # SCCM all
  $CPApplet = New-Object -Comobject CPApplet.CPAppletMgr
  $Actions = $CPApplet.GetClientActions()
  ForEach ($Action in $Actions) { $Action.PerformAction() } 
}

function Get-LastLoggedNonAdmin($pc) {
  $AdmDsk = @('drwin', 'Administrator', '58691', '10245', '53942')

  $opt = New-CimSessionOption -Protocol DCOM
  $s = New-CimSession -Computername $PC -SessionOption $opt -ErrorAction Stop

  $res = Get-CimInstance –ClassName Win32_UserProfile -Filter "Special = 'False' AND LastUseTime IS NOT NULL" -CimSession $s |
  Sort-Object -Property LastUseTime -Descending -Unique |
  Select LocalPath, LastUseTime, @{N = 'User'; E = { $_.LocalPath | % { $_.split('\')[-1] } } } -First 20  #| % {$_.split('\')[-1]} 
  #$res | ? { $_.user -notlike  "dsk_*" -and $_.user -notin $AdmDsk }
  $res | select LocalPath, LastUseTime, user, @{N = 'DN'; E = { (Get-ADUser $_.user -Properties DisplayName).DisplayName } }

  Remove-CimSession $s

  Get-CimSession | Remove-CimSession
}

function Copy-MaintanceWindow {
  $mw = Get-CMMaintenanceWindow -CollectionName "All Clients" 
  $mw | % { New-CMMaintenanceWindow -CollectionName "SCCM Group 4" -Name $_.Name -Schedule (Convert-CMSchedule -ScheduleString $_.ServiceWindowSchedules) -ApplyTo Any | select Name, Description, Duration }
}

function Deploy-File ($PCs, $File, $path = "C:\Temp\inst\", $run, $cmd) {
  $srcfile = split-path $file -Leaf
  # if (!$cmd)  { $cmd = "C:\Temp\inst\$srcfile" }
  $pcs | % {
    $destPath = "\\$_\" + ($path -replace ':', '$')
    if (-not (test-path "$destPath") ) { md $destPath -Verbose }
    if (-not (test-path "$destPath\$srcfile") ) { Copy-Item $file $destPath -Force -Verbose }
  }
}

function Get-PRDGroups([string]$uid) {
  $exist = [bool](Get-ADUser -Filter { SamAccountName -eq $uid } -Server prd.aib.pri) 
  if ($exist) {
    $usr = Get-ADUser -Identity $uid -Properties DisplayName -Server prd.aib.pri
    "PRD\$($usr.Name) - $($usr.DisplayName)"
    $all = Get-ADObject -Filter { Name -eq $usr.sid.Value } -Properties msds-principalname, memberof |  
    % { [PSCustomObject]@{ User = $_.'msds-principalname'; Group = ($_.memberof | Get-ADGroup).Name } }
    $all.Group | % { [PSCustomObject]@{ User = $all.User; Group = $_ } }
  }
  else { "$usr - User not found" }
}

function Get-DomainFromDist($dist) {
  ($dist -split ",DC=")[1]
}

function Get-GroupGroups($gr) {
  $grps = @(Get-ADGroupMember $gr | ? { $_.objectClass -eq 'group' } | ? { $_.Name -notin $temp })
  foreach ($g in $grps) {
    $global:temp += $g.Name
    Write-Progress -Activity "Processing $($g.name)" -Status "Retrieving data .." -PercentComplete (($grps.IndexOf($g) / $grps.Count) * 100)
    [PSCustomObject]@{ Domain = '>> GROUP'; User = $g.Name; DisplayName = ''; Description = '' }  
    Get-GroupUsers($g.SamAccountName) 
    Get-GroupGroups($g.SamAccountName) 
  }
}

function Get-GroupUsers($gr) {
  $users = Get-ADGroupMember $gr | ? { $_.objectClass -eq 'user' }
  foreach ($user in $users) {
    $domain = Get-DomainFromDist($user.distinguishedName)
    $uinfo = Get-ADUser $user -Properties DisplayName, Description -Server "$domain.aib.pri"
    [PSCustomObject]@{ Domain = $domain; User = $user.Name; DisplayName = $uinfo.displayName; Description = $uinfo.Description }
  }
}

function Get-GroupsAll($grp) {
  $gr = Get-ADGroup -Filter { Name -eq $grp } -Properties ManagedBy, Description
  $exist = [bool]$gr 
  if ($exist) {
    $global:temp = @()
    $all = @(Get-GroupUsers($gr)), (Get-GroupGroups($gr))
    Write-Progress "Processing " -Completed
    $currPath = if ($psISE) { Split-Path $psISE.CurrentFile.FullPath } else { $PSScriptRoot } 
    $filename = "$($grp) - $(get-date -Format 'yyyy-MM-dd HH-mm').txt"
    $manager = if ($gr.ManagedBy) { (get-aduser $gr.ManagedBy -Properties DisplayName).DisplayName } else { 'OWNER' }
    ""
    "$($gr.Name)`n$($manager)`n$($gr.Description)`n" + ($all | Out-String).TrimEnd() | tee $currPath\$filename
    "`nExported to file : ..\$filename"
  }
  else { "Group not found" } 
  "" 
}


Function Get-LastLoginInfo {

  <##requires -RunAsAdministrator
.Synopsis
    This will get a Information on the last users who logged into a machine.
    More info can be found: https://docs.microsoft.com/en-us/windows/security/threat-protection/auditing/basic-audit-logon-events
 
.NOTES
    Name: Get-LastLoginInfo
    Author: theSysadminChannel
    Version: 1.0
    DateCreated: 2020-Nov-27
 
.EXAMPLE
    Get-LastLoginInfo -ComputerName Server01, Server02, PC03 -SamAccountName username
 
.LINK
    https://thesysadminchannel.com/get-computer-last-login-information-using-powershell -
#>
 
  [CmdletBinding(DefaultParameterSetName = "Default")]
  param(
    [Parameter(
      Mandatory = $false,
      ValueFromPipeline = $true,
      ValueFromPipelineByPropertyName = $true,
      Position = 0
    )]
    [string[]]  $ComputerName = $env:COMPUTERNAME,
 
    [Parameter(
      Position = 1,
      Mandatory = $false,
      ParameterSetName = "Include"
    )]
    [string]    $SamAccountName,
 
    [Parameter(
      Position = 1,
      Mandatory = $false,
      ParameterSetName = "Exclude"
    )]
    [string]    $ExcludeSamAccountName,
 
    [Parameter(Mandatory = $false)]
    [ValidateSet("SuccessfulLogin", "FailedLogin", "Logoff", "DisconnectFromRDP")]
    [string]    $LoginEvent = "SuccessfulLogin",
 
    [Parameter(Mandatory = $false)] [int] $DaysFromToday = 3,
 
    [Parameter(Mandatory = $false)] [int] $MaxEvents = 1024,
 
    [System.Management.Automation.PSCredential] $Credential
  )
 
  BEGIN {
    $StartDate = (Get-Date).AddDays(-$DaysFromToday)
    Switch ($LoginEvent) {
      SuccessfulLogin { $EventID = 4624 }
      FailedLogin { $EventID = 4625 }
      Logoff { $EventID = 4647 }
      DisconnectFromRDP { $EventID = 4779 }
    }
  }
 
  PROCESS {
    foreach ($Computer in $ComputerName) {
      try {
        $Computer = $Computer.ToUpper()
        $Time = "{0:F0}" -f (New-TimeSpan -Start $StartDate -End (Get-Date) | Select -ExpandProperty TotalMilliseconds) -as [int64]
 
        if ($PSBoundParameters.ContainsKey("SamAccountName")) {
          $EventData = "
                        *[EventData[
                                Data[@Name='TargetUserName'] != 'SYSTEM' and
                                Data[@Name='TargetUserName'] != '$($Computer)$' and
                                Data[@Name='TargetUserName'] = '$($SamAccountName)'
                            ]
                        ]
                    "
        }
 
        if ($PSBoundParameters.ContainsKey("ExcludeSamAccountName")) {
          $EventData = "
                        *[EventData[
                                Data[@Name='TargetUserName'] != 'SYSTEM' and
                                Data[@Name='TargetUserName'] != '$($Computer)$' and
                                Data[@Name='TargetUserName'] != '$($ExcludeSamAccountName)'
                            ]
                        ]
                    "
        }
 
        if ((-not $PSBoundParameters.ContainsKey("SamAccountName")) -and (-not $PSBoundParameters.ContainsKey("ExcludeSamAccountName"))) {
          $EventData = "
                        *[EventData[
                                Data[@Name='TargetUserName'] != 'SYSTEM' and
                                Data[@Name='TargetUserName'] != '$($Computer)$'
                            ]
                        ]
                    "
        }
 
        $Filter = @"
                    <QueryList>
                        <Query Id="0">
                            <Select Path="Security">
                            *[System[
                                    Provider[@Name='Microsoft-Windows-Security-Auditing'] and
                                    EventID=$EventID and
                                    TimeCreated[timediff(@SystemTime) &lt;= $($Time)]
                                ]
                            ]
                            and
                                $EventData
                            </Select>
                        </Query>
                    </QueryList>
"@
 
        if ($PSBoundParameters.ContainsKey("Credential")) {
          $EventLogList = Get-WinEvent -ComputerName $Computer -FilterXml $Filter -Credential $Credential -ErrorAction Stop
        }
        else {
          $EventLogList = Get-WinEvent -ComputerName $Computer -FilterXml $Filter -ErrorAction Stop
        }
 
 
        $Output = foreach ($Log in $EventLogList) {
          #Removing seconds and milliseconds from timestamp as this is allow duplicate entries to be displayed
          $TimeStamp = $Log.timeCReated.ToString('MM/dd/yyyy hh:mm tt') -as [DateTime]
 
          switch ($Log.Properties[8].Value) {
            2 { $LoginType = 'Interactive' }
            3 { $LoginType = 'Network' }
            4 { $LoginType = 'Batch' }
            5 { $LoginType = 'Service' }
            7 { $LoginType = 'Unlock' }
            8 { $LoginType = 'NetworkCleartext' }
            9 { $LoginType = 'NewCredentials' }
            10 { $LoginType = 'RemoteInteractive' }
            11 { $LoginType = 'CachedInteractive' }
          }
 
          if ($LoginEvent -eq 'FailedLogin') {
            $LoginType = 'FailedLogin'
          }
 
          if ($LoginEvent -eq 'DisconnectFromRDP') {
            $LoginType = 'DisconnectFromRDP'
          }
 
          if ($LoginEvent -eq 'Logoff') {
            $LoginType = 'Logoff'
            $UserName = $Log.Properties[1].Value.toLower()
          }
          else {
            $UserName = $Log.Properties[5].Value.toLower()
          }

          [PSCustomObject]@{
            ComputerName = $Computer
            TimeStamp    = $TimeStamp
            UserName     = $UserName
            LoginType    = $LoginType
          }
        }
 
        #Because of duplicate items, we'll append another select object to grab only unique objects
        $Output | select ComputerName, TimeStamp, UserName, LoginType -Unique | select -First $MaxEvents
 
      }
      catch { Write-Error $_.Exception.Message }
    }
  }
  END {}
}

function Download-Edge {

  <#
.SYNOPSIS
  Get-EdgeEnterpriseMSI

.DESCRIPTION
  Imports all device configurations in a folder to a specified tenant

.PARAMETER Channel
  Channel to download, Valid Options are: Dev, Beta, Stable, EdgeUpdate, Policy.

.PARAMETER Platform
  Platform to download, Valid Options are: Windows or MacOS, if using channel "Policy" this should be set to "any"
  Defaults to Windows if not set.

.PARAMETER Architecture
  Architecture to download, Valid Options are: x86, x64, arm64, if using channel "Policy" this should be set to "any"
  Defaults to x64 if not set.

.PARAMETER Version
  If set the script will try and download a specific version. If not set it will download the latest.

.PARAMETER Folder
  Specifies the Download folder

.PARAMETER Force
  Overwrites the file without asking.

.NOTES
  Version:        1.2
  Author:         Mattias Benninge
  Creation Date:  2020-07-01

  Version history:

  1.0 -   Initial script development
  1.1 -   Fixes and improvements by @KarlGrindon
          - Script now handles multiple files for e.g. MacOS Edge files
          - Better error handling and formating
          - URI Validation
  1.2 -   Better compability on servers (force TLS and remove dependency to IE)

  
  https://docs.microsoft.com/en-us/mem/configmgr/apps/deploy-use/deploy-edge

.EXAMPLE
  
  Download the latest version for the Beta channel and overwrite any existing file
  .\Get-EdgeEnterpriseMSI.ps1 -Channel Beta -Folder D:\SourceCode\PowerShell\Div -Force

#>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $false, HelpMessage = 'Channel to download, Valid Options are: Dev, Beta, Stable, EdgeUpdate, Policy')]
    [ValidateSet('Dev', 'Beta', 'Stable', 'EdgeUpdate', 'Policy')]
    [string]$Channel = 'Stable',
  
    [Parameter(Mandatory = $False, HelpMessage = 'Folder where the file will be downloaded')]
    [ValidateNotNullOrEmpty()]
    [string]$Folder = 'c:\Temp',

    [Parameter(Mandatory = $false, HelpMessage = 'Platform to download, Valid Options are: Windows or MacOS')]
    [ValidateSet('Windows', 'MacOS', 'any')]
    [string]$Platform = "Windows",

    [Parameter(Mandatory = $false, HelpMessage = "Architecture to download, Valid Options are: x86, x64, arm64, any")]
    [ValidateSet('x86', 'x64', 'arm64', 'any')]
    [string]$Architecture = "x64",

    [parameter(Mandatory = $false, HelpMessage = "Specifies which version to download")]
    [ValidateNotNullOrEmpty()]
    [string]$ProductVersion,

    [parameter(Mandatory = $false, HelpMessage = "Overwrites the file without asking")]
    [Switch]$Force
  )

  $ErrorActionPreference = "Stop"
  $edgeEnterpriseMSIUri = 'https://edgeupdates.microsoft.com/api/products?view=enterprise'

  # Validating parameters to reduce user errors
  if ($Channel -eq "Policy" -and ($Architecture -ne "Any" -or $Platform -ne "Any")) {
    Write-Warning ("Channel 'Policy' requested, but either 'Architecture' and/or 'Platform' is not set to 'Any'. 
                  Setting Architecture and Platform to 'Any'")
    $Architecture = "Any"
    $Platform = "Any"
  } 
  elseif ($Channel -ne "Policy" -and ($Architecture -eq "Any" -or $Platform -eq "Any")) {
    throw "If Channel isn't set to policy, architecture and/or platform can't be set to 'Any'"
  }
  elseif ($Channel -eq "EdgeUpdate" -and ($Architecture -ne "x86" -or $Platform -eq "Windows")) {
    Write-Warning ("Channel 'EdgeUpdate' requested, but either 'Architecture' is not set to x86 and/or 'Platform' 
                  is not set to 'Windows'. Setting Architecture to 'x64' and Platform to 'Windows'")
    $Architecture = "x64"
    $Platform = "Windows"
  }

  #Write-Host "Enabling connection over TLS for better compability on servers" -ForegroundColor Green
  [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

  # Test if HTTP status code 200 is returned from URI
  try { Invoke-WebRequest $edgeEnterpriseMSIUri -UseBasicParsing | Where-Object StatusCode -match 200 | Out-Null }
  catch { throw "Unable to get HTTP status code 200 from $edgeEnterpriseMSIUri. Does the URL still exist?" }
  Write-Host "Getting available files from $edgeEnterpriseMSIUri" -ForegroundColor Green

  # Try to get JSON data from Microsoft
  try {
    $response = Invoke-WebRequest -Uri $edgeEnterpriseMSIUri -Method Get -ContentType "application/json" -UseBasicParsing -ErrorVariable InvokeWebRequestError
    $jsonObj = ConvertFrom-Json $([String]::new($response.Content))
    Write-Host "Succefully retrived data" -ForegroundColor Green
  }
  catch { throw "Could not get MSI data: $InvokeWebRequestError" }

  # Alternative is to use Invoke-RestMethod to get a Json object directly
  # $jsonObj = Invoke-RestMethod -Uri "https://edgeupdates.microsoft.com/api/products?view=enterprise" -UseBasicParsing

  $selectedIndex = [array]::indexof($jsonObj.Product, "$Channel")

  if (-not $ProductVersion) {
    try {
      Write-host "No version specified, getting the latest for $Channel" -ForegroundColor Green
      $selectedVersion = (([Version[]](($jsonObj[$selectedIndex].Releases |
              Where-Object { $_.Architecture -eq $Architecture -and $_.Platform -eq $Platform }).ProductVersion) |
          Sort-Object -Descending)[0]).ToString(4) 
      Write-Host "Latest Version for channel $Channel is $selectedVersion`n" -ForegroundColor Green
      $selectedObject = $jsonObj[$selectedIndex].Releases |
      Where-Object { $_.Architecture -eq $Architecture -and $_.Platform -eq $Platform -and $_.ProductVersion -eq $selectedVersion }
    }
    catch { throw "Unable to get object from Microsoft. Check your parameters and refer to script help." }
  }
  else {
    Write-Host "Matching $ProductVersion on channel $Channel" -ForegroundColor Green
    $selectedObject = ($jsonObj[$selectedIndex].Releases |
      Where-Object { $_.Architecture -eq $Architecture -and $_.Platform -eq $Platform -and $_.ProductVersion -eq $ProductVersion })
    if (-not $selectedObject) {
      throw "No version matching $ProductVersion found in $channel channel for $Architecture architecture."
    }
    else { Write-Host "Found matching version`n" -ForegroundColor Green }
  }


  if (Test-Path $Folder) {
    foreach ($artifacts in $selectedObject.Artifacts) {
      # Not showing the progress bar in Invoke-WebRequest is quite a bit faster than default
      $ProgressPreference = 'SilentlyContinue'   
      Write-host "Starting download of: $($artifacts.Location)" -ForegroundColor Green
      # Work out file name
      $fileName = Split-Path $artifacts.Location -Leaf
      if (Test-Path "$Folder\$fileName" -ErrorAction SilentlyContinue) {
        if ($Force) {
          Write-Host "Force specified. Will attempt to download and overwrite existing file." -ForegroundColor Green
          try { Invoke-WebRequest -Uri $artifacts.Location -OutFile "$Folder\$fileName" -UseBasicParsing }
          catch { throw "Attempted to download file, but failed: $error[0]" }    
        }
        else {
          # CR-someday: There should be an evaluation of the file version, if possible. Currently the function only
          # checks if a file of the same name exists, not if the versions differ
          Write-Host "$Folder\$fileName already exists!" -ForegroundColor Yellow
          do { $overWrite = Read-Host -Prompt "Press Y to overwrite or N to quit." }
          # -notmatch is case insensitive
          while ($overWrite -notmatch '^y$|^n$')
          if ($overWrite -match '^y$') {
            Write-Host "Starting Download" -ForegroundColor Green
            try { Invoke-WebRequest -Uri $artifacts.Location -OutFile "$Folder\$fileName" -UseBasicParsing }
            catch { throw "Attempted to download file, but failed: $error[0]" }
          }
          else {
            Write-Host "File already exists and user chose not to overwrite, exiting script." -ForegroundColor Red
            exit
          }
        }
      }
      else {
        Write-Host "Starting Download" -ForegroundColor Green
        try { Invoke-WebRequest -Uri $artifacts.Location -OutFile "$Folder\$fileName" -UseBasicParsing }
        catch { throw "Attempted to download file, but failed: $error[0]" }
      }
      if (((Get-FileHash -Algorithm $artifacts.HashAlgorithm -Path "$Folder\$fileName").Hash) -eq $artifacts.Hash) {
        Write-Host "Calculated checksum matches known checksum`n" -ForegroundColor Green
      }
      else {
        Write-Warning "Checksum mismatch!"
        Write-Warning "Expected Hash: $($artifacts.Hash)"
        Write-Warning "Downloaded file Hash: $((Get-FileHash -Algorithm $($artifacts.HashAlgorithm) -Path "$Folder\$fileName").Hash)`n"
      }
    }
  }
  else { throw "Folder $Folder does not exist" }
  Write-Host "-- Script Completed: File Downloaded -- " -ForegroundColor Green
}

Function Set-Window {
  <#
.SYNOPSIS
Retrieve/Set the window size and coordinates of a process window.

.DESCRIPTION
Retrieve/Set the size (height,width) and coordinates (x,y) 
of a process window.

.PARAMETER ProcessName
Name of the process to determine the window characteristics. 
(All processes if omitted).

.PARAMETER Id
Id of the process to determine the window characteristics. 

.PARAMETER X
Set the position of the window in pixels from the left.

.PARAMETER Y
Set the position of the window in pixels from the top.

.PARAMETER Width
Set the width of the window.

.PARAMETER Height
Set the height of the window.

.PARAMETER Passthru
Returns the output object of the window.

.NOTES
Name:   Set-Window
Author: Boe Prox
Version History:
    1.0//Boe Prox - 11/24/2015 - Initial build
    1.1//JosefZ   - 19.05.2018 - Treats more process instances 
                                 of supplied process name properly
    1.2//JosefZ   - 21.02.2019 - Parameter Id

.OUTPUTS
None
System.Management.Automation.PSCustomObject
System.Object

.EXAMPLE
Get-Process powershell | Set-Window -X 20 -Y 40 -Passthru -Verbose
VERBOSE: powershell (Id=11140, Handle=132410)

Id          : 11140
ProcessName : powershell
Size        : 1134,781
TopLeft     : 20,40
BottomRight : 1154,821

Description: Set the coordinates on the window for the process PowerShell.exe

.EXAMPLE
$windowArray = Set-Window -Passthru
WARNING: cmd (1096) is minimized! Coordinates will not be accurate.

    PS C:\>$windowArray | Format-Table -AutoSize

  Id ProcessName    Size     TopLeft       BottomRight  
  -- -----------    ----     -------       -----------  
1096 cmd            199,34   -32000,-32000 -31801,-31966
4088 explorer       1280,50  0,974         1280,1024    
6880 powershell     1280,974 0,0           1280,974     

Description: Get the coordinates of all visible windows and save them into the
             $windowArray variable. Then, display them in a table view.

.EXAMPLE
Set-Window -Id $PID -Passthru | Format-Table
​‌‍
  Id ProcessName Size     TopLeft BottomRight
  -- ----------- ----     ------- -----------
7840 pwsh        1024,638 0,0     1024,638

Description: Display the coordinates of the window for the current 
             PowerShell session in a table view.
             

     
#>
  [cmdletbinding(DefaultParameterSetName = 'Name')]
  Param (
    [parameter(Mandatory = $False,
      ValueFromPipelineByPropertyName = $True, ParameterSetName = 'Name')]
    [string]$ProcessName = '*',
    [parameter(Mandatory = $True,
      ValueFromPipeline = $False, ParameterSetName = 'Id')]
    [int]$Id,
    [int]$X,
    [int]$Y,
    [int]$Width,
    [int]$Height,
    [switch]$Passthru
  )
  Begin {
    Try { 
      [void][Window]
    }
    Catch {
      Add-Type @"
        using System;
        using System.Runtime.InteropServices;
        public class Window {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetWindowRect(
            IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public extern static bool MoveWindow(  
            IntPtr handle, int x, int y, int width, int height, bool redraw);
              
        [DllImport("user32.dll")] 
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ShowWindow(
            IntPtr handle, int state);
        }
        public struct RECT
        {
        public int Left;        // x position of upper-left corner
        public int Top;         // y position of upper-left corner
        public int Right;       // x position of lower-right corner
        public int Bottom;      // y position of lower-right corner
        }
"@
    }
  }
  Process {
    $Rectangle = New-Object RECT
    If ( $PSBoundParameters.ContainsKey('Id') ) {
      $Processes = Get-Process -Id $Id -ErrorAction SilentlyContinue
    }
    else {
      $Processes = Get-Process -Name "$ProcessName" -ErrorAction SilentlyContinue
    }
    if ( $null -eq $Processes ) {
      If ( $PSBoundParameters['Passthru'] ) { 
        Write-Warning 'No process match criteria specified'
      }
    }
    else {
      $Processes | ForEach-Object {
        $Handle = $_.MainWindowHandle
        Write-Verbose "$($_.ProcessName) `(Id=$($_.Id), Handle=$Handle`)"
        if ( $Handle -eq [System.IntPtr]::Zero ) { return }
        $Return = [Window]::GetWindowRect($Handle, [ref]$Rectangle)
        If (-NOT $PSBoundParameters.ContainsKey('X')) {
          $X = $Rectangle.Left            
        }
        If (-NOT $PSBoundParameters.ContainsKey('Y')) {
          $Y = $Rectangle.Top
        }
        If (-NOT $PSBoundParameters.ContainsKey('Width')) {
          $Width = $Rectangle.Right - $Rectangle.Left
        }
        If (-NOT $PSBoundParameters.ContainsKey('Height')) {
          $Height = $Rectangle.Bottom - $Rectangle.Top
        }
        If ( $Return ) {
          $Return = [Window]::MoveWindow($Handle, $x, $y, $Width, $Height, $True)
        }
        If ( $PSBoundParameters['Passthru'] ) {
          $Rectangle = New-Object RECT
          $Return = [Window]::GetWindowRect($Handle, [ref]$Rectangle)
          If ( $Return ) {
            $Height = $Rectangle.Bottom - $Rectangle.Top
            $Width = $Rectangle.Right - $Rectangle.Left
            $Size = New-Object System.Management.Automation.Host.Size        -ArgumentList $Width, $Height
            $TopLeft = New-Object System.Management.Automation.Host.Coordinates -ArgumentList $Rectangle.Left , $Rectangle.Top
            $BottomRight = New-Object System.Management.Automation.Host.Coordinates -ArgumentList $Rectangle.Right, $Rectangle.Bottom
            If ($Rectangle.Top -lt 0 -AND 
              $Rectangle.Bottom -lt 0 -AND
              $Rectangle.Left -lt 0 -AND
              $Rectangle.Right -lt 0) {
              Write-Warning "$($_.ProcessName) `($($_.Id)`) is minimized! Coordinates will not be accurate."
            }
            $Object = [PSCustomObject]@{
              Id          = $_.Id
              ProcessName = $_.ProcessName
              Size        = $Size
              TopLeft     = $TopLeft
              BottomRight = $BottomRight
            }
            $Object
          }
        }
      }
    }
  }
}

Function Get-ScreenColor {
  <#
    .SYNOPSIS
    Gets the color of the pixel under the mouse, or of the specified space.
    .DESCRIPTION
    Returns the pixel color either under the mouse, or of a location onscreen using X/Y locating.  If no parameters are supplied, the mouse cursor position will be retrived and used.

    Current Version - 1.0
    .EXAMPLE
    Mouse-Color
    Returns the color of the pixel directly under the mouse cursor.
    .EXAMPLE
    Mouse-Color -X 300 -Y 300
    Returns the color of the pixel 300 pixels from the top of the screen and 300 pixels from the left.
    .PARAMETER X
    Distance from the top of the screen to retrieve color, in pixels.
    .PARAMETER Y
    Distance from the left of the screen to retrieve color, in pixels.
    .NOTES

    Revision History
    Version 1.0
        - Live release.  Contains two parameter sets - an empty default, and an X/Y set.
    #>

  #Requires -Version 4.0

  [CmdletBinding(DefaultParameterSetName = 'None')]

  param(
    [Parameter(
      Mandatory = $true,
      ParameterSetName = "Pos"
    )]
    [Int]
    $X,
    [Parameter(
      Mandatory = $true,
      ParameterSetName = "Pos"
    )]
    [Int]
    $Y
  )
    
  if ($PSCmdlet.ParameterSetName -eq 'None') {
    $pos = [System.Windows.Forms.Cursor]::Position
  }
  else {
    $pos = New-Object psobject
    $pos | Add-Member -MemberType NoteProperty -Name "X" -Value $X
    $pos | Add-Member -MemberType NoteProperty -Name "Y" -Value $Y
  }
  $map = [System.Drawing.Rectangle]::FromLTRB($pos.X, $pos.Y, $pos.X + 1, $pos.Y + 1)
  $bmp = New-Object System.Drawing.Bitmap(1, 1)
  $graphics = [System.Drawing.Graphics]::FromImage($bmp)
  $graphics.CopyFromScreen($map.Location, [System.Drawing.Point]::Empty, $map.Size)
  $pixel = $bmp.GetPixel(0, 0)
  $red = $pixel.R
  $green = $pixel.G
  $blue = $pixel.B
  $result = New-Object psobject
  if ($PSCmdlet.ParameterSetName -eq 'None') {
    $result | Add-Member -MemberType NoteProperty -Name "X" -Value $([System.Windows.Forms.Cursor]::Position).X
    $result | Add-Member -MemberType NoteProperty -Name "Y" -Value $([System.Windows.Forms.Cursor]::Position).Y
  }
  $result | Add-Member -MemberType NoteProperty -Name "Red" -Value $red
  $result | Add-Member -MemberType NoteProperty -Name "Green" -Value $green
  $result | Add-Member -MemberType NoteProperty -Name "Blue" -Value $blue
  return $result
}

<#


# getting first domain component value from distinguishedName
$user.DistinguishedName -replace '^.*?DC=|,DC=.*$'

# constructing domain FQDN from DistinguishedName
$user.DistinguishedName -replace '^.*?,DC=' -replace ',DC=','.'

# extracting domain FQDN from canonicalname
$user.CanonicalName -replace '/.*$'

# Getting first Domain component value from canonicalname
$user.CanonicalName -replace '\..*$'

List a Cm Group
(Get-CMCollectionMember -CollectionName 'Group 2').name | % { [PSCustomObject]@{PC=$_; Desc=(Get-ADComputer $_ -Properties description).description}  }

Add computer to Collection
(Get-CMCollectionMember -CollectionName 'Office 2016 group 5').name | % { Add-CMDeviceCollectionDirectMembershipRule -CollectionName “Group 2” -ResourceID (Get-CMDevice -Name $_).ResourceID }

SCCM
new collection
3..10 | % { $NewCol = New-CMDeviceCollection -Name “Group $($_)” -LimitingCollectionName “All Clients” -RefreshType Both 
Move-CMObject -FolderPath “.\DeviceCollection\22H2 Upgrade” -InputObject $NewCol}



# get MAC Address
# Solution 1
Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter "IPEnabled='True'" -ComputerName 3R6DG52-DUB | 
Select-Object -Property MACAddress, Description
 
# Solution 2
Get-WmiObject -ClassName Win32_NetworkAdapterConfiguration -Filter "IPEnabled='True'" -ComputerName $pc | 
Select-Object -Property MACAddress, Description
 

 # taskkill /s 6PN6MM2-DUB /fi "IMAGENAME eq excel*"
# set-itemproperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings' -name ProxyEnable -value 1
#  'ss s s    s s ' -replace '\s+', ' '



working -  $returnval = ([WMICLASS]"\\W10-MB\ROOT\CIMV2:win32_process").Create("C:\Temp\jre-8u311-windows-i586.exe `/s")

([WMICLASS]"\\7V0TGL2-BCS\ROOT\CIMV2:win32_process").Create("\\W10-mb\c$\Temp\jre-8u311-windows-i586.exe `/s")


$staging.Name | % { 
 ADD-ADGroupMember "BCM Deployment Group Win 10" –members "$_$" -Verbose
 $ou = (Get-ADOrganizationalUnit -Filter { name -like "Treasury Win 10 PC*" }).DistinguishedName
 get-adcomputer $_ | Set-ADComputer -Description "CP Build (on bench)" -PassThru -Verbose | Move-ADObject -TargetPath "$ou" -Verbose
}

iex ${using:function:Test-Modules}.Ast.Extent.Text;Test-Modules


(Get-ADPrincipalGroupMembership (Get-ADComputer w10-mb).DistinguishedName).name | ? { $_ -like "*deploy*"}

#>
