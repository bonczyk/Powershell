# Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1"
# Import-Module "\\drsitsrv1\DRSsupport$\Projects\2022\Test-BCS\modules\MBMod\0.3\MBMod.psm1" -Force -WarningAction SilentlyContinue
# Import-Module "H:\MB\PS\modules\MBMod\0.3\MBMod.psm1" -Force -WarningAction SilentlyContinue
# Import-Module "\\fxt8\c$\H\MB\PS\modules\MBMod\0.3\MBMod.psm1" -Force -WarningAction SilentlyContinue
# [Management.Automation.WildcardPattern]::Escape('test[1].txt')
# [regex]::Escape("(foo)")
# Run-Remote $list "powershell -command `"Start-Transcript c:\temp\rb_log12-18-22.txt; Get-appxprovisionedpackage –online -Verbose | where-object {`$_.displayname -like \`"*Edge*\`" }; Stop-Transcript`""
# test-path C:\WINDOWS\System32\DriverStore\FileRepository\hpcu250v.inf_amd64_f9bc7c093f784e4e\hpcu250v.inf

function Test-Modules {
  Init
  $path = "$ScriptPath\modules"
  if ($ScriptPath -eq $ModulePath2) {$path = Split-Path (Split-Path $ScriptPath)}

  $modUNC = @{ ImportExcel = "$path\ImportExcel\7.4.1\ImportExcel.psd1"
   InvokePsExec = "$path\InvokePsExec\1.2\InvokePsExec.psd1"
                     #MBMod = "$path\MBMod\0.3\MBMod.psm1"
             }
  $ModUNC.keys.ForEach( { If (-not(Get-module $_)) { Import-Module $($ModUNC[$_]) -Global -WA SilentlyContinue } })
}

function ImportMe {
  #iex ${using:function:ImportMe}.Ast.Extent.Text;ImportMe
  Import-Module "H:\MB\PS\modules\MBMod\0.3\MBMod.psm1" -WA SilentlyContinue
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

function Init {
  #$ErrorActionPreference='silentlycontinue'
  $global:ModulePath  = 'H:\MB\PS\modules\MBMod\0.3\' 
  $global:ModulePath1 = $PSCommandPath
  $global:ModulePath2 = (Get-Module -Name mbmod).ModuleBase
  $global:ScriptFile  =  Get-CallingFileName
  $global:ScriptPath  = if ($x = Get-CallingFileName) { Split-Path $x } else { ScriptDir }
  $global:DesktopPath = [Environment]::GetFolderPath("Desktop")
  $global:PatternSID  = 'S-1-5-21-\d+-\d+\-\d+\-\d+$'

  $global:upath = "$ModulePath\users.xlsx"
  $global:cpath = "$ModulePath\comps.xlsx"
  #"Mß v0.9.9"
}

function Main {

}

function Get-FileDetails($path) {
$objShell = New-Object -ComObject Shell.Application 
$objFolder = $objShell.namespace((Get-Item $path).DirectoryName) 

foreach ($File in $objFolder.items()) {
    IF ($file.path -eq $path) {
        $FileMetaData = New-Object PSOBJECT 
        for ($a=0 ; $a -le 266; $a++) {  
         if($objFolder.getDetailsOf($File, $a)) { 
             $hash += @{$($objFolder.getDetailsOf($objFolder.items, $a)) = $($objFolder.getDetailsOf($File, $a)) }
            $FileMetaData | Add-Member $hash 
            $hash.clear()  
           } 
       }
    }
}
return $FileMetaData
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
 [PSCustomObject]@{ pc=$pc; ver = $ver; build = $build }
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
  if ($text) { "$text - $(Get-Date -Format 'yyyy-MM-dd HH-mm')" }
  else { "$(get-date -Format 'yyyy-MM-dd HH-mm')" }
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
  set-itemproperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings' -name ProxyEnable -value $val
}

function Set-WorkWeekSchedule($ProgramName,$CollectionName,$time) {
 #Get-CMPackageDeployment -ProgramName $ProgramName | Select-Object PackageID -ExpandProperty AssignedSchedule 
 $a = 1..5 | % { New-CMSchedule -DayOfWeek $_ -Start (Get-Date -F "dd/MM/yy $time") }
 Get-CMDeployment -ProgramName $ProgramName -CollectionName $CollectionName | Set-CMPackageDeployment -StandardProgramName $ProgramName -Schedule $a  
}


function Run-Remote($Pc,$Cmd,$Timeout=3,$CurrentDir=’C:\temp’) {
 $opt = New-CimSessionOption -Protocol DCOM
  try {
    $s = New-CimSession -Computername $pc -SessionOption $opt -OperationTimeoutSec $timeout -ErrorAction Stop  
    Invoke-CimMethod Win32_Process -method Create @{CommandLine="cmd /c $cmd"; CurrentDirectory=$CurrentDir} -CimSession $s
    Remove-CimSession $s 
  } catch { $false } 
}

# usage : Run-Remote w10-mb "dir nosuchfile.txt > c:\temp\mm.txt 2>&1"


function Run-Remote_WMIold($pc,$cmd) {
 ([WMICLASS]"\\$pc\ROOT\CIMV2:win32_process").Create($cmd).ProcessId
}

function Check-WMI($pc,$timeout=3) {
  $opt = New-CimSessionOption -Protocol DCOM
  try {
    $s = New-CimSession -Computername $pc -SessionOption $opt -OperationTimeoutSec $timeout -ErrorAction Stop 
    $t = (get-date) - (gcim Win32_OperatingSystem -CimSession $s -ErrorAction SilentlyContinue).LastBootUpTime
    Remove-CimSession $s 
    [bool]$t
  } catch { $false } 
}

function Get-MissingDrivers($pc) {
#For formatting:
    $result = @{Expression = {$_.Name}; Label = "Device Name"},
              @{Expression = {$_.ConfigManagerErrorCode} ; Label = "Status Code" }

#Checks for devices whose ConfigManagerErrorCode value is greater than 0, i.e has a problem device.
Get-WmiObject -Class Win32_PnpEntity -ComputerName $pc -Namespace Root\CIMV2 | Where-Object {$_.ConfigManagerErrorCode -gt 0 } | select name,ConfigManagerErrorCode #| Format-Table $result -AutoSize
}

function MyProgress ($text, $maxcount) { #rv ii -ErrorAction SilentlyContinue
  If (-not(Test-Path Variable:\ii)) { $global:ii = 0 }
  $global:ii++
  If ($global:ii -gt $maxcount) { $global:ii = 0 } 
  $perc = [math]::Round($ii / $maxcount * 100, 1);
  Write-Progress $text "Complete : $perc %" -perc $perc
}

function Menu ($Title, [array]$opt) {
  "$Title"
  '-'*20
  for ($i = 1; $i -lt $opt.count + 1; $i++) { 
    if ($opt[$i-1] -eq 'back to search') 
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

function hl ($text, $word, $fc, $bc) {
  $text = ($text | Out-String).Trim()
  $s = $text -split $word
  if (!$fc) { $fc = 14 }
  Write-Host $s[0] -NoNewline
  for ($i = 1; $i -lt $s.count; $i++) {  
    $ex = "Write-Host $word -NoNewline -ForegroundColor $fc "
    if ($bc) { $ex += "-BackgroundColor Yellow" } 
    iex $ex
    Write-Host $s[$i] -NoNewline
  }
}

Function WinTitle($Title) {
  $host.ui.RawUI.WindowTitle = $Title
}

function RemoveUserProfile($PC,$user){
 $opt = New-CimSessionOption -Protocol DCOM
 $s = New-CimSession -Computername $PC -SessionOption $opt -ErrorAction Stop
 #Get-CimInstance -Class Win32_UserProfile -CimSession $S | SELECT LocalPath
 Get-CimInstance -Class Win32_UserProfile -CimSession $S | Where-Object { $_.LocalPath.split('\')[-1] -eq $user } | Remove-CimInstance
 Remove-CimSession $s
}

function ADinfo {
  #Write-Debug "Updating AD from servers"
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
  $prop = @('msDS-UserPasswordExpiryTimeComputed', 'Name', 'DisplayName', 'Description', 'Office', 'mail', 'LastBadPasswordAttempt', 'BadPwdCount', 'LockedOut', 'pwdLastSet')
  #$global:ADu = Get-ADUser -Filter * -Properties $prop | ? { $_.name -match '^\d{5}$' } 
  $tempU.AddRange( (Get-ADUser -Filter * -Properties $prop ) )  # | ? { $_.name -match '^\d{5}$' }
  $tempU | % {
    $val = if ($_.'msDS-UserPasswordExpiryTimeComputed' -eq '9223372036854775807') { 'Password Never Expired' }
    else { Get-Date ([DateTime]::FromFileTime([Int64]::Parse($_.'msDS-UserPasswordExpiryTimeComputed'))) } # -Format "dd/MM/yyyy HH:mm:ss"   ([datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")) 
    $_ | Add-Member -MemberType NoteProperty -Name 'ExpiryDate' -Value $val -Force
    $_ | Add-Member -MemberType NoteProperty -Name 'LastPwdSet' -Value (Get-Date ([DateTime]::FromFileTime([Int64]::Parse($_.pwdLastSet)))) -Force
    if ($_.pwdLastSet) { $_.pwdLastSet  }

  } 
  $ex = 'msDS-UserPasswordExpiryTimeComputed', 'pwdLastSet', 'WriteDebugStream', 'WriteErrorStream', 'WriteInformationStream', 'WriteVerboseStream', 'WriteWarningStream', 'PropertyNames', 'AddedProperties', 'RemovedProperties', 'ModifiedProperties', 'PropertyCount'
  $ADu.AddRange( ($tempU | select * -ExcludeProperty $ex) )
  Remove-Variable TempU
  $r = 'St Helens', 'London', '1st Floor,', ' 1 Undershaft', 'Old Jewry' -join '|'
  $ADu | % { $_.office = ($_.office -replace $r).trim() } #($r|%{if ($u -match $_){$u -replace $_,''}}).Trim()
}

function Get-DealersPCs {
  $global:ADc = New-Object System.Collections.Generic.List[System.Object]
  $TempC = New-Object System.Collections.Generic.List[System.Object]
  $TempC.AddRange( (Get-ADComputer -Filter { OperatingSystem -NotLike "*server*" } -prop description,location) )
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
  $DCs.AddRange( (Get-ADDomainController -Filter * | select -Skip 1) )
  $DCs.AddRange( (Get-ADDomainController -Filter * -Server prd.aib.pri | Select -First 10) )
  $online = APing($DCs.hostname)
  Foreach ($DC in $online) {
    $t = Get-ADUser -Identity $user -Server $DC.Name -Properties AccountLockoutTime, LastBadPasswordAttempt, BadPwdCount, LockedOut, pwdLastSet, msDS-UserPasswordExpiryTimeComputed
    if ($t) {
      Add-Member -InputObject $t -MemberType NoteProperty -Name DC -Value $DC.Name -Force
      Add-Member -InputObject $t -MemberType NoteProperty -Name LastPwdSet -Value (Get-Date ([DateTime]::FromFileTime([Int64]::Parse($t.pwdLastSet)))) -Force
      Add-Member -InputObject $t -MemberType NoteProperty -Name ExpiryTime -Value (Get-Date ([DateTime]::FromFileTime([Int64]::Parse($t.'msDS-UserPasswordExpiryTimeComputed')))) -Force 
    }
    else { $dc.name }
    $t | Select DC, Name, Enabled, LockedOut,@{N='LastBad'; E={$_.LastBadPasswordAttempt} }, @{N='BadCount'; E={$_.BadPwdCount} }, LastPwdSet, ExpiryTime
  }  
}

function Set-Console($title, $width, $height) {
  if ($title) { $host.UI.RawUI.WindowTitle = $Title }
  if ($width) { [console]::WindowWidth = $width; [console]::BufferWidth = [console]::WindowWidth }
  if ($height) { [console]::WindowHeight = $height }
}

function SetWinTitle($p,$text) {
if ("Win32Api" -as [type]) {} else {
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
  
public static class Win32Api
{
    [DllImport("User32.dll", EntryPoint = "SetWindowText")]
    public static extern int SetWindowText(IntPtr hWnd, string text);
}
"@}
 # How to use 
 #$p = Start-Process -FilePath "notepad.exe" -PassThru
 #$p.WaitForInputIdle() | out-null #only GUI
 [Win32Api]::SetWindowText($p.MainWindowHandle, $text)  
}

Function Execute-Command ($commandTitle, $commandPath, $commandArguments){
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
        stdout = $p.StandardOutput.ReadToEnd()
        stderr = $p.StandardError.ReadToEnd()
        ExitCode = $p.ExitCode
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

function New-PSWin-Alert($in) { #do not use security alerts
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

  if ($pc) { $pc | Set-Clipboard; "`n'$pc' has been copied to the clipboard`n" }
  Menu "Choose option" @('Show Lockout Status', 'Unlock', "Go to $pc", 'New console window', 'Back')
  '' 
  $inp = Read-Host "[1-5] "
  switch ($inp) {
    '1' { Receive-Job -Name 'LockoutStatus' -Wait | select * -ExcludeProperty RunspaceId,PSSourceJobInstanceId | ft}
    '2' { "Admin rights needed to unlock account - Unlock-ADAccount $u" }
    '3' { Check-PC $pc }       # "Set-ADAccountPassword $u -Reset -NewPassword (ConvertTo-SecureString -AsPlainText 'p@ssw0rd' -Force) " }
    '4' { New-PSWin $user }
    '0' { "back to search" }
    Default { "back to search" }
  }
}

function New-DameWare($pc) {
  $dw = "C:\Program Files (x86)\DameWare Development\DameWare Mini Remote Control\DWRCC.exe"
  $cmd = "-m:$pc -a:1" # -h -c"
  #iex  "&'$dw' $cmd"
  start "$dw" "$cmd"     
}

function New-MSTSC($PCs) {
  $PCs | % { mstsc /v:$_ }
}

function New-PingWindow($ip) {
  start-process cmd -ArgumentList "/C","mode con:cols=55 lines=10 && title Ping $ip && powershell -command ""&{(get-host).ui.rawui.buffersize=@{width=55;height=200};}"" && ping $ip -t"
}

function Get-ADRealUsers {
Get-ADUser -Filter { Surname -like "*" -and memberof -like '*'  } -prop name,givenname,surname `
 | select name,givenname,surname # | Export-Excel -Path C:\Users\dsk_58691\Desktop\usr.xlsx
}

function Check-PC($pc) {
    $p = $ADc | ? { $_.Name -eq $pc }
    $on = APing($pc)
    $l = Get-LoggedUsers; [array]$LLast = $LoggedUsers | ? { $_.Computer -eq $pc }
    "`n"*2
    ($p | select Name,description,DNSHostName | fl | Out-String ).Trim()  
    if ($on) { 
        $uptime = Get-BootTimeF $pc
        $LNow = Logged-User $pc    
        "Online      : $($on.Address)" 
        "Up Time     : $uptime"
        "Logged User : $($LNow.USERNAME)  $($LNow.DisplayName)  $($LNow.'LOGON TIME')  $($LNow.SESSIONNAME)"
        
    } else { "Offline !! "};''

    $Opt = @( "Open C: - \\$pc\c$",
              "Open comand prompt on $pc",
              "DameWare $pc",
              "Remote Desktop $pc"
              'New PS console window',
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
      '5' { New-PSWin $pc }
      '6' { WOL $pc; New-PingWindow($pc) }
      '7' { New-PingWindow($pc) }
      '8' { Restart-Computer $pc -Force; New-PingWindow($pc) }
      '9' { compmgmt.msc -a /computer=$pc }
      '0' { "back to search" }
      Default { "back to search" }
    }
}

function Get-GraphicDrivers($pc) {
Get-WmiObject Win32_VideoController -ComputerName $pc | ForEach-Object {
      [PSCustomObject]@{
        ComputerName  = $_.SystemName
        Description   = $_.Description -join ', '
        DriverDate    = [DateTime]::ParseExact($_.DriverDate -replace '000000.000000-000', 'yyyyMMdd', $culture).ToString('yyyy-MM-dd')
        DriverVersion = $_.DriverVersion
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
    Init;Get-ADinfo
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
    $LoggedUsers = ($LoggedUsers | ? { $_.Username -ne 'NONE' -and $_.displayName}) ### !!!!
  }

  # if LoggedUsers not exist
  & $sb_import
}

function Logged-User {
  [CmdletBinding()]Param([Parameter(ValueFromPipeline)]$pc)
  process {
    
    if ($pc -eq "") { $pc = $env:COMPUTERNAME }
    $o = [PScustomObject]@{ Computer = $pc; Description = ($Adc | ? { $_.name -eq $pc }).Description; 
      USERNAME = ''; DisplayName = ''; SESSIONNAME = ''; ID = ''; STATE = ''; 'IDLE TIME' = ''; 'LOGON TIME' = '';
      dt = (get-date -Format G)    
    }
    if ($pc -ne '8P1PJ32-BCS') { 
      if (APing $pc) {
        try {
          $temp = (query user /server:$pc 2>&1)  
          If ($temp) { # If ($temp -split '`n' -eq 'No User exists for *') {$temp = $null; $user = $false}
            $r = $temp -replace '\s{2,}', ',' | ConvertFrom-Csv
            $r.psobject.Properties.name | % { $o.$_ = $r.$_ }
            $o.DisplayName = ($adu | ? { $_.name -eq $r.USERNAME }).DisplayName 
          }
        }
        catch { $o.USERNAME = 'NONE' }
      }
      else { $o.USERNAME = 'OFFLINE' }
    }
    $o
  }
}

function isLogged($pc = "$env:COMPUTERNAME") {
$i = 0; $user = $null; $r = $null
#if ($pc -eq '8P1PJ32-BCS') { continue }
if (APing($pc)) {
  try {
    $temp = (query user /server:$pc 2>&1)  
    If ($temp -split '`n' -eq 'No User exists for *') {$temp = $null; $user = $false}
    If ($temp) { 
      $r = $temp -replace '\s{2,}', ',' | ConvertFrom-Csv 
      If ($r.USERNAME[0] -eq '>') { $r.USERNAME = $r.username.Substring(1) }
      $User = $r.USERNAME }
  } catch { $user = 'error' }
} else { $user = 'pcoff' }
return [PSCustomObject]@{ PC = $pc; User = $user }
} 




function Get-BootTime ($pc) {
  $opt = New-CimSessionOption -Protocol DCOM
  try {
    $s = New-CimSession -Computername $pc -SessionOption $opt -OperationTimeoutSec 3 -ErrorAction Stop
    $t = (get-date) - (gcim Win32_OperatingSystem -CimSession $s -ErrorAction SilentlyContinue).LastBootUpTime
    Remove-CimSession $s }
  catch { $t = 0 }
    
  [PScustomObject]@{ PC = $pc; up = $t; }
}

function Get-BootTimeF ($pc) {
  (Get-BootTime $pc).up.tostring("dd\.hh\:mm\:ss")
}

function Get-UnusedCN($uptime=72) {
  Write-Progress "Getting list of unused computers" "..." -perc 0
  $l = (Get-ADComputer -Filter {OperatingSystem -NotLike "*server*" }).name # Write-verbose "Getting list of computers from AD where OperatingSystem is not like *server*" -and Name -like "*-DUB"
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
  $togo | % { if ( (isLogged $_.pc).user -eq $false ) { $_;Restart-Computer $_.pc -Force -Verbose -ErrorAction SilentlyContinue; $restarted+=,$_.pc } }
  $restarted | Out-File "$ScriptPath\$(sdate RestartLog).txt" -Append #Stop-Transcript
}

function Check-Logs {
  # calculate start time (one hour before now)
  $Start = (Get-Date) - (New-Timespan -Hours 1)
  $Computername = $env:COMPUTERNAME 
 
  # Getting all event logs
  Get-EventLog -AsString -ComputerName $Computername |
  ForEach-Object {
    # write status info
    Write-Progress -Activity "Checking Eventlogs on \\$ComputerName" -Status $_

    # get event entries and add the name of the log this came from
    Get-EventLog -LogName $_ -EntryType Error, Warning -After $Start -ComputerName $ComputerName -ErrorAction SilentlyContinue |
    Add-Member NoteProperty EventLog $_ -PassThru 
       
  } |
  # sort descending
  Sort-Object -Property TimeGenerated -Descending |
  # select the properties for the report
  Select-Object EventLog, TimeGenerated, EntryType, Source, Message | 
  # output into grid view window
  Out-GridView -Title "All Errors & Warnings from \\$Computername"
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
    $SiteServer = "drscmsrv1.dealers.aib.pri"
)
 
Write-Verbose "CmpName = $CmpName"
Write-Verbose "CollId  = $CollID"
Write-Verbose "SiteServer = $SiteServer"

if (!$CmpName -and !$CollId) { Write-Warning "Please provide ComputerName or CollectionID to WOL" ; break }
if (!$CmpName -and $CollId -eq "SMS00001") {
    Write-Warning "Seems wrong to wake every single computer in the environment, refusing to perform." ; break  }
 
$SiteCode = (Get-WmiObject -ComputerName "$SiteServer" -Namespace root\sms -Query 'SELECT SiteCode FROM SMS_ProviderLocation').SiteCode
 
if ($CmpName) {
  $ResourceID = (Get-WmiObject  -ComputerName "$SiteServer" -Namespace "Root\SMS\Site_$($SiteCode)" -Query "Select ResourceID from SMS_R_System Where NetBiosName = '$($CmpName)'").ResourceID
  if ($ResourceID) { $CmpName = @($ResourceID) }
}
 
$WMIConnection = [WMICLASS]"\\$SiteServer\Root\SMS\Site_$($SiteCode):SMS_SleepServer"
$Params = $WMIConnection.psbase.GetMethodParameters("MachinesToWakeup")
$Params.MachineIDs = $CmpName
$Params.CollectionID  = $CollId
$return = $WMIConnection.psbase.InvokeMethod("MachinesToWakeup", $Params, $Null) 
 
if (!$return) {
  Write-Host "No machines are online to wake up selected devices" }
if ($return.numsleepers -ge 1) {
  Write-Host "The resource selected are scheduled to wake-up as soon as possible" } 
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
  ($(foreach ($bp in $Global:MyInvocation.BoundParameters.GetEnumerator()) { # argument list
      $valRep =
        if ($bp.Value -is [switch]) { # switch parameter
          if ($bp.Value) { $sep = '' } # switch parameter name by itself is enough
          else { $sep = ':'; '$false' } # `-switch:$false` required
        }
        else { # Other data types, possibly *arrays* of values.
          $sep = ' '
          foreach ($val in $bp.Value) {
            if ($val -is [bool]) { # a Boolean parameter (rare)
              ('$false', '$true')[$val] # Booleans must be represented this way.
            } else { # all other types: stringify in a culture-invariant manner.
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

function Admin {  #[environment]::GetCommandLineArgs()
 Init
 pushd "$ScriptPath"
 if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { 
  Start-Process powershell.exe -Verb RunAs "-NoProfile -ExecutionPolicy Bypass -File `"$(Get-CallingFileName)`" $(Get-BoundParam) $(Get-UnboundParam)" ; exit }
 popd
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

function RemotePopup($pc,$text) {
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

function Get-Displays($pc) {
  (Get-WmiObject -ComputerName $pc win32_VideoController).name
  Get-WmiObject -ComputerName $pc WmiMonitorID -Namespace root\wmi | Select @{n = "Connected To"; e = { ($_.__Server) } }, @{n = "Make_Model"; e = { [System.Text.Encoding]::ASCII.GetString($_.UserFriendlyName -ne 00) } }, @{n = "Serial Number"; e = { [System.Text.Encoding]::ASCII.GetString($_.SerialNumberID -ne 00) } } | Out-GridView
}

function Accelerators {
  $TAType = [psobject].Assembly.GetType("System.Management.Automation.TypeAccelerators")
  $TAType::Add('accelerators', $TAType)

  [accelerators]::Get   # this now works
}

function Send-Enter {
  $pinvokes = @'
  [DllImport("user32.dll", CharSet=CharSet.Auto)]
  public static extern IntPtr FindWindow(IntPtr sClassName, string lpWindowName);
  [DllImport("user32.dll")]
  [return: MarshalAs(UnmanagedType.Bool)]
  public static extern bool SetForegroundWindow(IntPtr hWnd);
'@
  Add-Type -AssemblyName System.Windows.Forms
  Add-Type -MemberDefinition $pinvokes -Name My -Namespace MB

  while ([MB.My]::FindWindow([intptr]::zero, "McAfee Agent") -eq 0) {
    sleep -Milliseconds 300
  }
  $hwnd = [MB.My]::FindWindow([intptr]::zero, "McAfee Agent")
  if ($hwnd) {
    [MB.My]::SetForegroundWindow($hwnd)
    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
  }
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

function Get-InstalledApp2 {
[cmdletbinding()]            
param(            
 [parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]            
 [string[]]$ComputerName, #(Get-Content list.txt),       #$env:computername,   
 [String[]]$Name
)            
            
begin {   
 if (-not $ComputerName) { if (-not (Test-path list.txt)) { $ComputerName = (Get-ADComputer -Filter {OperatingSystem -NotLike "*server*"}).name } else { $ComputerName = Get-Content list.txt } }
 $UninstallRegKeys=@("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall",            
     "SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall")           
}            
            
process { 
 $i = 0          
 foreach($Computer in $ComputerName) {  
   $perc = [math]::Round($i/$ComputerName.Count*100,1)
   Write-Progress "Getting Information from computer $Computer" "Complete : $perc %" -perc $perc;  $i++   
   Write-Verbose "Working on $Computer"            
 if(Test-Connection -ComputerName $Computer -Count 1 -ea 0) {            
  foreach($UninstallRegKey in $UninstallRegKeys) {            
   try {            
    $HKLM   = [microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$computer)            
    $UninstallRef  = $HKLM.OpenSubKey($UninstallRegKey)            
    $Applications = $UninstallRef.GetSubKeyNames()            
   } catch {            
    Write-Verbose "Failed to read $UninstallRegKey"            
    Continue            
   }            
            
   foreach ($App in $Applications) {     
     foreach ($Nam in $Name) {   
   $AppRegistryKey  = $UninstallRegKey + "\\" + $App            
   $AppDetails   = $HKLM.OpenSubKey($AppRegistryKey)            
   $AppGUID   = $App            
   $AppDisplayName  = $($AppDetails.GetValue("DisplayName"))  
   if ($AppDisplayName -notlike $Nam)  { continue }
   $AppVersion   = $($AppDetails.GetValue("DisplayVersion"))            
   $AppPublisher  = $($AppDetails.GetValue("Publisher"))            
   $AppInstalledDate = $($AppDetails.GetValue("InstallDate"))            
   $AppUninstall  = $($AppDetails.GetValue("UninstallString"))            
   if($UninstallRegKey -match "Wow6432Node") {            
    $Softwarearchitecture = "x86" } else { $Softwarearchitecture = "x64" }            
   if(!$AppDisplayName) { continue }            
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
   $all += ,$OutputObj 
   }
   }            
  }             
 }  else {
      $OutputObj = New-Object -TypeName PSobject             
      $OutputObj | Add-Member -MemberType NoteProperty -Name ComputerName -Value $Computer.ToUpper()            
      $OutputObj | Add-Member -MemberType NoteProperty -Name AppName -Value "OFFLINE" 
      $OutputObj 
      $all += ,$OutputObj
    }     
 }            
}            
            
end {}
}


function Get-InstalledApp {
[cmdletbinding()]            
param(            
 [parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]            
 [string[]]$ComputerName, #(Get-Content list.txt),       #$env:computername,   
 [String[]]$Name
)            
            
begin {   
 if (-not $ComputerName) { if (-not (Test-path list.txt)) { $ComputerName = (Get-ADComputer -Filter {OperatingSystem -NotLike "*server*"}).name } else { $ComputerName = Get-Content list.txt } }
 $UninstallRegKeys=@("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall",            
     "SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall")           
}            
            
process { 
 $i = 0          
 foreach($Computer in $ComputerName) {  
   $perc = [math]::Round($i/$ComputerName.Count*100,1)
   Write-Progress "Getting Information from computer $Computer" "Complete : $perc %" -perc $perc;  $i++   
   Write-Verbose "Working on $Computer"            
 if(Aping $Computer) {            
  foreach($UninstallRegKey in $UninstallRegKeys) {            
   try {        
    $HKLM   = [microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$computer)            
    $UninstallRef  = $HKLM.OpenSubKey($UninstallRegKey)            
    $Applications = $UninstallRef.GetSubKeyNames()            
   } catch { Write-Verbose "Failed to read $UninstallRegKey"; Continue }            
            
   foreach ($App in $Applications) {     
     foreach ($Nam in $Name) {   
       $AppRegistryKey  = $UninstallRegKey + "\\" + $App            
       $AppDetails   = $HKLM.OpenSubKey($AppRegistryKey)                       
       $AppDisplayName  = $($AppDetails.GetValue("DisplayName"))  
       if (!$AppDisplayName -or $AppDisplayName -notlike $Nam)  { continue }                         
       [PSCustomObject]@{
             ComputerName = $Computer.ToUpper();
                  AppName = $AppDisplayName;
               AppVersion = $AppDetails.GetValue("DisplayVersion");
                AppVendor = $AppDetails.GetValue("Publisher");
            InstalledDate = $AppDetails.GetValue("InstallDate");
             UninstallKey = $AppDetails.GetValue("UninstallString");
                  AppGUID = $AppGUID = $App;
     SoftwareArchitecture = if($UninstallRegKey -match "Wow6432Node") {"x86"} else { "x64" }    
  }}}}             
 } else { [PSCustomObject]@{ ComputerName = $Computer.ToUpper(); AppName = 'OFFLINE';} }     
}}                      
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
    [int]$ConsoleCount = 3
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
  Set-PSReadLineKeyHandler -Key ctrl+B -BriefDescription 'show busy' -LongDescription "make it look like I am working" -ScriptBlock {
    param($key, $arg)
    #Add-Type -Assembly PresentationCore
    [Microsoft.PowerShell.PSConsoleReadLine]::RevertLine();
    [Microsoft.PowerShell.PSConsoleReadLine]::Insert('Show-MeBeingSuperBusy -ConsoleCount 3; clear;');
    [Microsoft.PowerShell.PSConsoleReadLine]::AcceptLine();
  }
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
  try { $user = gwmi -Class win32_computersystem -ComputerName $computer | select -ExpandProperty username -ErrorAction Stop } 
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
  [void][Threading.Tasks.Task]::WaitAll($Task,200) 
  $Task.Where( { $_.result.status -eq 'success' }) | % { $_.result | Add-Member -NotePropertyName Name -NotePropertyValue $_.name -Force -ErrorAction SilentlyContinue; $_.result | select * -ExcludeProperty RoundtripTime,Options,Buffer} 
}

function APing2($PCs) {
  $buffer = ([system.text.encoding]::ASCII).getbytes("a" * [int]32)
  $Task = ForEach ($PC in $PCs) {
    (New-Object System.Net.NetworkInformation.Ping).SendPingAsync($PC, 200, $buffer, @{TTL = 128; DontFragment = $false }) | Add-Member -NotePropertyName Name -NotePropertyValue $pc -PassThru -Force 
  } 
  [void][Threading.Tasks.Task]::WaitAll($Task,200) 
  $Task | % { $_.result | Add-Member -NotePropertyName Name -NotePropertyValue $_.name -Force -ErrorAction SilentlyContinue; $_.result | select * -ExcludeProperty RoundtripTime,Options,Buffer} 
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
    $inp = Read-Host -Prompt $txt
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

} finally {
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

function Uninstall-Wmi{
[cmdletbinding()]            
param (            
 [parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
 [string]$ComputerName = $env:computername,
 [parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,Mandatory=$true)]
 [string]$AppGUID
)            

 try {
  $returnval = ([WMICLASS]"\\$computerName\ROOT\CIMV2:win32_process").Create("msiexec `/x$AppGUID `/norestart `/qn")
 } catch {
  write-error "Failed to trigger the uninstallation. Review the error message"
  $_
 }
 switch ($($returnval.returnvalue)){
  0 { "Uninstallation command triggered successfully" }
  2 { "You don't have sufficient permissions to trigger the command on $Computer" }
  3 { "You don't have sufficient permissions to trigger the command on $Computer" }
  8 { "An unknown error has occurred" }
  9 { "Path Not Found" }
  9 { "Invalid Parameter"}
 }
 }

function uninst-java {
$list = (Get-ADComputer -Filter {OperatingSystem -NotLike "*server*"}).name #(Get-ADComputer -Filter {Name -like "*-bcs"} -SearchBase "OU=CTS Win 10 PC``s,OU=DRS Win 10 PCs,DC=dealers,DC=aib,DC=pri").name
$on = (aping($list)).name
rv ii,all -ErrorAction SilentlyContinue
[System.Collections.ArrayList]$all = @()
$all = Get-InstalledApp $c "*java 8*"
Export-Xlsx -obj $all -path 'C:\Users\dsk_58691\Desktop\uninst-java-all.xlsx'
}

function UpdateGraphicDrivers($pc,$drvPath) {

# Import-Module "H:\MB\PS\modules\MBMod\0.1\MBMod.psm1" -Force -WarningAction SilentlyContinue
# check for NVIDIA drivers
# $pc | % { Logged-User $_ } | ft
# $pc | % { Get-GraphicDrivers $_ | ? { $_.Description -like "*NVIDIA*"} } | sort ComputerName -Unique | sort DriverDate


$srcfile = split-path $drvPath -Leaf
$c = "C:\Temp\inst\" + $srcfile + " -s -n Display.Driver"
$x=0;$out=@()

$pc | % { 
  $destPath = "\\$_\c$\Temp\inst\"
  if (-not (test-path "$destPath") ) { md $destPath -Verbose}
  if (-not (test-path "$destPath\$srcfile") ) { Copy-Item $srcPath $destPath -Force -Verbose}
  [PSCustomObject]@{ PC = $_ ; PID = (Run-Remote $_ $c) }
}

}

function Get-CMCollectionOfDevice {
    [CmdletBinding()]
    [OutputType([int])]
    Param
    (
        # Computername
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [String]$Computer,
 
        # ConfigMgr SiteCode
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [String]$SiteCode = "DRS",
 
        # ConfigMgr SiteServer
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   Position=2)]
        [String]$SiteServer = "drscmsrv1.dealers.aib.pri"
    )
Begin
{
    [string] $Namespace = "root\SMS\site_$SiteCode"
}
 
Process
{
    $si=1
    Write-Progress -Activity "Retrieving ResourceID for computer $computer" -Status "Retrieving data" 
    $ResIDQuery = Get-WmiObject -ComputerName $SiteServer -Namespace $Namespace -Class "SMS_R_SYSTEM" -Filter "Name='$Computer'"
    
    If ([string]::IsNullOrEmpty($ResIDQuery))
    {
        Write-Output "System $Computer does not exist in Site $SiteCode"
    }
    Else
    {
    $Collections = (Get-WmiObject -ComputerName $SiteServer -Class sms_fullcollectionmembership -Namespace $Namespace -Filter "ResourceID = '$($ResIDQuery.ResourceId)'")
    $colcount = $Collections.Count
    
    $devicecollections = @()
    ForEach ($res in $collections)
    {
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
 
End
{
    $devicecollections
}
}

function Speak($text) {
  Add-Type -AssemblyName System.speech
  $speak = New-Object System.Speech.Synthesis.SpeechSynthesizer
  $speak.Rate = 3
  $speak.Speak($text) 
}

function Get-PCgroup($pc){
  (Get-ADPrincipalGroupMembership (Get-ADComputer $pc).DistinguishedName).name 
}


function SCCM-ForceUpd($pc){
$strAction = "{00000000-0000-0000-0000-000000000121}" # Application Deployment Evaluation Cycle
try
    {
        $WMIPath = "\\" + $pc + "\root\ccm:SMS_Client" 
        $SMSwmi = [wmiclass] $WMIPath 
        [Void]$SMSwmi.TriggerSchedule($strAction)
    }
catch
    { $_.Exception.Message }  
}


function Find7050IntelDriver {
$list = (Get-ADComputer -Filter {OperatingSystem -NotLike "*server*"}).name
 foreach ($pc in $list) {
  $opt = New-CimSessionOption -Protocol DCOM
    try {
  $s = New-CimSession -Computername $pc -SessionOption $opt -ErrorAction Stop -OperationTimeoutSec 2
  $model = (Get-CimInstance Win32_ComputerSystem -CimSession $s -Property Model).model
  if ($model -like "*7050") {
    Get-CimInstance Win32_PnPSignedDriver -Filter 'DeviceName LIKE "Intel(R) Chipset SATA%"' -CimSession $s | % { [PSCustomObject]@{ CN = $pc; DriverVer = $_.DriverVersion; DriverDate = $_.DriverDate; DeviceName = $_.devicename;} }
  }
  Remove-CimSession $s
    } catch { } 
 }
 Export-Excel -Path "$env:USERPROFILE\Desktop\RAPID.xlsx" -InputObject $all
}


function Replace-Links($pc,$chromelnk) {
 $path1 = "\\$pc\c$\ProgramData\Microsoft\Windows\Start Menu\Programs\Google Chrome.lnk"
 if (compare (gc $chromelnk) (gc $path1)) { Copy-Item -Path $chromelnk -Destination (Split-Path $path1) -Force -Verbose} else { "Correct - $path1" }
 $userlist = (Get-ChildItem "\\$pc\c$\Users\" -Directory -Exclude Administrator,drwin,public,default*).Name 
 $userlist | % { 
   $p = "\\$pc\c$\Users\$_\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Google Chrome.lnk"  
   if (Test-Path $p) { if (compare (gc $chromelnk) (gc $p)) { Copy-Item -Path $chromelnk -Destination (Split-Path $p) -Force -Verbose} else { "Correct - $p" }  }
 }
}


function WOL-IP {
$Mac = "D8:9E:F3:13:5C:7B"
$MacByteArray = $Mac -split "[:-]" | ForEach-Object { [Byte] "0x$_"}
[Byte[]] $MagicPacket = (,0xFF * 6) + ($MacByteArray  * 16)
$UdpClient = New-Object System.Net.Sockets.UdpClient
$UdpClient.Connect(([System.Net.IPAddress]::Parse('10.28.222.14')),7)
$UdpClient.Send($MagicPacket,$MagicPacket.Length)
$UdpClient.Close()
}



function Get-Wsus($ServerName='drsopsmgr2') {
  [reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration") | out-null
  [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer($ServerName,$false,8530) 
}

Function GetUpdateState {
param([string[]]$kbnumber='KB5016616',[string]$wsusserver='drsopsmgr2',[string]$port=8530
)
$report = @()
[void][reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration")
$wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer($wsusserver,$False,8530)
$CompSc = new-object Microsoft.UpdateServices.Administration.ComputerTargetScope
$updateScope = new-object Microsoft.UpdateServices.Administration.UpdateScope; 
$updateScope.UpdateApprovalActions = [Microsoft.UpdateServices.Administration.UpdateApprovalActions]::Install
foreach ($kb in $kbnumber){ #Loop against each KB number passed to the GetUpdateState function 
   $updates = $wsus.GetUpdates($updateScope) | ?{$_.Title -match $kb} #Getting every update where the title matches the $kbnumber
       foreach($update in $updates){ #Loop against the list of updates I stored in $updates in the previous step
          $update.GetUpdateInstallationInfoPerComputerTarget($CompSc) | ?{$_.UpdateApprovalAction -eq "Install"} |  % { #for the current update
#Getting the list of computer object IDs where this update is supposed to be installed ($_.UpdateApprovalAction -eq "Install")
          $Comp = $wsus.GetComputerTarget($_.ComputerTargetId)# using #Computer object ID to retrieve the computer object properties (Name, #IP address)
          $info = "" | select UpdateTitle, LegacyName, SecurityBulletins, Computername, OS ,IpAddress, UpdateInstallationStatus, UpdateApprovalAction #Creating a custom PowerShell object to store the information
          $info.UpdateTitle = $update.Title
          $info.LegacyName = $update.LegacyName
          $info.SecurityBulletins = ($update.SecurityBulletins -join ';')
          $info.Computername = $Comp.FullDomainName
          $info.OS = $Comp.OSDescription
          $info.IpAddress = $Comp.IPAddress
          $info.UpdateInstallationStatus = $_.UpdateInstallationState
          $info.UpdateApprovalAction = $_.UpdateApprovalAction
          $report+=$info # Storing the information into the $report variable 
        }
     }
  }
$report | ?{$_.UpdateInstallationStatus -ne 'NotApplicable' -and $_.UpdateInstallationStatus -ne 'Unknown' -and $_.UpdateInstallationStatus -ne 'Installed' } #|  Export-Csv -Path c:\temp\rep_wsus.csv -Append -NoTypeInformation #Filtering the report to list only computers where the updates are not installed
} # Usage: GetUpdateState -kbnumber KB5016616 -wsusserver drsopsmgr2 -port 8530

Init;Main

function Get-PcInfo {
[cmdletbinding()]
param(
    [parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
    [string[]]$ComputerName = (Read-Host -Prompt 'Please enter computer name') )
$ErrorActionPreference='silentlycontinue'

$Apps = "Adobe Acrobat Reader DC","Citrix online plug-in","Symantec_EnterpriseVault","PhishMe Reporter","Google Chrome",
        "Java 8 Update","Skype for Business 2016","Microsoft Office Standard 2013","QlikView Plugin","WinZip_","","McAfee Agent","McAfee Endpoint","Tanium"

$hostn = $ComputerName                
$user  = $env:username                 #(Get-WmiObject -Class Win32_ComputerSystem | Select-Object UserName).Username.Split('\')[1]
$file  = "H:\Builds\ToDo\${hostn}.txt"

function showsave($text) {
 $text
 $text >> $file
}

$name=(Get-WmiObject Win32_OperatingSystem -ComputerName $ComputerName).caption      #Microsoft Windows 7\10 Enterprise
$bit=(Get-WmiObject Win32_OperatingSystem -ComputerName $ComputerName).OSArchitecture
$ver=0;  #(Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ReleaseId
$build = (gwmi Win32_OperatingSystem -ComputerName $ComputerName).Version 
if ($build -eq '10.0.18362') { $ver = '19H1' } 
if ($build -eq '10.0.18363') { $ver = '19H2' } 
if ($build -eq '10.0.19041') { $ver = '20H1' }
if ($build -eq '10.0.19042') { $ver = '20H2' } 
if ($build -eq '10.0.19043') { $ver = '21H1' } 
if ($build -eq '10.0.19044') { $ver = '21H2' }

$czas = (Get-Date).ToString('MM/dd/yyyy hh:mm:ss tt')
$dn   = ([adsisearcher]"(&(objectClass=user)(samaccountname=$user))").FindOne().Properties['displayname']
$ip   = (Test-Connection -ComputerName $computername -count 1).IPV4Address.ipaddressTOstring
$vid  = @(Get-WmiObject Win32_VideoController -ComputerName $ComputerName) | ? { $_.name -ne 'DameWare Development Mirror Driver 64-bit' -and $_.name -ne 'Microsoft Remote Display Adapter' }
if (@($vid).count -gt 1) { $vid="$($vid[0].name) + $($vid[1].name)"} else { $vid=$vid.name }

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
  $Monitors=@(Get-WmiObject win32_desktopmonitor);  
  showsave("MonitorNo : $($Monitors.count)`n") 
 } 

#$tmp = $(Get-PSDrive -PSProvider FileSystem | Where-Object {$_.DisplayRoot -ne $null} | select Name,DisplayRoot | ft -hidetableheaders)
#$tmp.Count
#showsave($tmp)

function numInstances([string]$process) {
    @(Get-Process $process -ErrorAction 0).Count
}

$Array = @()
Foreach ($app in $Apps) {
 $Result=[PSCustomObject]@{ Name = $app; IsIns = if ($app) {if ( (Get-InstalledApp $ComputerName "*$app*" | ? { $_.appName -ne 'OFFLINE' } | measure).count -ne 0 ) {$true} else {$false} } }
 $Array += $Result
}
showsave(($Array | Format-Table -HideTableHeaders | Out-String).Trim())
showsave("Tanium process no `t`t: " + $(numInstances("TaniumClient")))

}


function WordFill {

 $template = 'G:\Inventory\DRS Desktop Build & Decommission signoffs\Windows 10 Build Checklist Template.docx'
 $wf='C:\Temp\alloc\Windows 10 Build Sheet.docx'
 $fold = 'H:\Builds\ToDo'
 $done = 'H:\Builds\DoneByMe'
 $file = gci $fold *.txt | select -First 1 
 $fn = $file.FullName
 $fn

 function RemoveColon ($fn,$nr) {
   $line = (Get-Content $fn)[$nr]
   $start = $line.IndexOf(':') + 1
   $result = $line.Substring($start,$line.Length - $start).Trim()
   return $result
 }

 if ( !(Test-Path (Split-Path $wf)) ) { mkdir (Split-Path $wf) | Out-Null }
 if ( !(Test-Path ($wf)) ) { copy $template $wf }

 $l = Get-Content $fn -TotalCount 2  # (Get-Content $fn)[2]
 $time = $l.Replace('-','').Replace('=','').Trim()

 $wd = New-Object -ComObject Word.Application 
 $wd.Visible = $fasle
 $Doc = $Wd.Documents.Open($wf)
 #$Doc = $wd.Documents.Open($wordf, $false, $true)
 #$Sel = $Wd.Selection # $sel.StartOf(15)  $sel.MoveDown()

 $t1=$wd.ActiveDocument.Tables.item(1)
 $t1.Cell(2,1).Range.Text=RemoveColon $fn 4
 $t1.Cell(2,2).Range.Text=RemoveColon $fn 8
 $t1.Cell(2,3).Range.Text=RemoveColon $fn 10
 $t1.Cell(2,4).Range.Text=RemoveColon $fn 11
 $t1.Cell(2,5).Range.Text=RemoveColon $fn 12
 $t1.Cell(2,6).Range.Text=(RemoveColon $fn 13) + "`n" + (RemoveColon $fn 14)
# $t1.Cell(4,1).Range.Text="Old Hostname"
 
 $t2=$wd.ActiveDocument.Tables.item(2)
 $t2.Cell(2,2).Range.Text=(RemoveColon $fn 3).Split('-').Trim()[0] #(RemoveColon $fn 3).substring(0,5)
 $t2.Cell(2,1).Range.Text=(RemoveColon $fn 3).Split('-').Trim()[1] 
 
 $t3=$wd.ActiveDocument.Tables.item(3)
 for ($i = 0; $i -lt 11; $i++) { 
   $t3.Cell(3+$i,2).Range.Text= (Get-Content $fn)[16+$i] -split " " | ? { $_ } | select -Last 1 #next tanium and mcaffee
 }
 $t3.Cell(2+$i,2).Range.Text=(Get-Content $fn)[27] -split " " | ? { $_ } | select -Last 1 
 $t3.Cell(3+$i,2).Range.Text=(Get-Content $fn)[29] -split " " | ? { $_ } | select -Last 1 

 for ($i = 0; $i -lt 6; $i++) { 
   $t3.Cell(38+$i,2).Range.Text="Done"
 }

 $t4=$wd.ActiveDocument.Tables.item(4)
 $t4.Cell(2,3).Range.Text = (Get-Content $fn)[28] -split " " | ? { $_ } | select -Last 1 
 for ($i = 0; $i -lt 8; $i++) { 
  $t4.Cell(3+$i,3).Range.Text="Done"
 }

 $t5=$wd.ActiveDocument.Tables.item(5)
 $t5.Cell(1,2).Range.text="Maciej Bonczyk"
 $t5.Cell(1,3).Range.text=(get-date).ToString("dd/MM/yyyy")

 $saveas = Join-Path $fold -ChildPath ((RemoveColon $fn 4)+'.docx')
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



<#

# get MAC Address
# Solution 1
Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter "IPEnabled='True'" -ComputerName 3R6DG52-DUB | 
Select-Object -Property MACAddress, Description
 
# Solution 2
Get-WmiObject -ClassName Win32_NetworkAdapterConfiguration -Filter "IPEnabled='True'" -ComputerName 3R6DG52-DUB | 
Select-Object -Property MACAddress, Description
 

 # taskkill /s 6PN6MM2-DUB /fi "IMAGENAME eq excel*"
# set-itemproperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings' -name ProxyEnable -value 1
#  'ss s s    s s ' -replace '\s+', ' '



working -  $returnval = ([WMICLASS]"\\W10-MB\ROOT\CIMV2:win32_process").Create("C:\Temp\jre-8u311-windows-i586.exe `/s")

([WMICLASS]"\\7V0TGL2-BCS\ROOT\CIMV2:win32_process").Create("\\W10-mb\c$\Temp\jre-8u311-windows-i586.exe `/s")

"\\drstreassrv2.dealers.aib.pri\Droom\JR TEST\infos\MapDrive\jre-8u311-windows-i586.exe"

5DBT762-DUB,C164KF2-DUB,6PP4MM2-DUB,6PN6MM2-DUB,7W7X65J-DUB,24XMS62-DUB,H79W65J-DUB,254LS62-DUB,AIBTL-3M6WG62,AIBTL-4M5465J,6PP3MM2-DUB,6NS5MM2-DUB,4W7X65J-DUB,FBBT762-DUB,W10-LOUISA,6NQ8MM2-DUB,CBCK992-DUB,24TRS62-DUB,AIBTL-7D1465J,CBGP992-DUB,24GMS62-DUB,AIBTL-4WGXG62,C
CMJF4J-DUB,4CFSG62-DUB,3MKTG62-BEL,3MDXG62-BEL,3MNCF62-BEL,GV7X65J-DUB,6VP4G4J-DUB,717FLG2-DUB,AIBTL-4SVVG62,259MS62-DUB,JWN4422-DUB,W10-MB,4CGCF62-DUB,4CFCF62-DUB,6NT2MM2-BCS,W10-ALYSSON,6FQ44K3-DUB,D5Z71K3-BCS,C5Z71K3-BCS,1FQ44K3-BCS,CDQ44K3-BCS,BDQ44K3-BCS,8FQ44K3-BCS,
4GQ44K3-BCS,3GQ44K3-BCS,FFQ44K3-DUB,8GQ44K3-BCS,6GQ44K3-BCS,7TZXGL2-BCS,6NV2MM2-BCS,7V0TGL2-BCS

  if ( @($x | ? { $_.AppVersion -ne '8.0.3110.11' }).count -eq 1) 
  {
    $x = $x | ? { $_.AppVersion -ne '8.0.3110.11' }
    if ($x) 
      { $x | % { 
          $o = [PSCustomObject]@{ PC=$l; newest=$True; Version=$_.AppVersion }  
          $o;  [void]$all.Add($o)
          #Uninstall-Wmi -ComputerName $l -AppGUID $_.AppGUID;  
          Export-Excel -Path 'C:\Users\dsk_58691\Desktop\uninst-java.xlsx' -InputObject $o -Append
        } 
      }



$staging.Name | % { 
 ADD-ADGroupMember "BCM Deployment Group Win 10" –members "$_$" -Verbose
 $ou = (Get-ADOrganizationalUnit -Filter { name -like "Treasury Win 10 PC*" }).DistinguishedName
 get-adcomputer $_ | Set-ADComputer -Description "CP Build (on bench)" -PassThru -Verbose | Move-ADObject -TargetPath "$ou" -Verbose
}

iex ${using:function:Test-Modules}.Ast.Extent.Text;Test-Modules

Delete user profile

$CN = "W10-MB"

$opt = New-CimSessionOption -Protocol DCOM
$s = New-CimSession -Computername $cn -SessionOption $opt -ErrorAction Stop

Get-CimInstance -Class Win32_UserProfile -CimSession $S | SELECT LocalPath

Get-CimInstance -Class Win32_UserProfile -CimSession $S | Where-Object { $_.LocalPath.split('\')[-1] -eq 'dsk_53942' } | Remove-CimInstance

Remove-CimSession $s

(Get-ADPrincipalGroupMembership (Get-ADComputer w10-mb).DistinguishedName).name | ? { $_ -like "*deploy*"}

#>
