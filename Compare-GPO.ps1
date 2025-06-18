function Compare-GPO {
  param ([String]$Gpo1='22H2C-V1 Bank & Country Credit',[String]$Gpo2='24H2C-V1 WIN 11 B&CC')
  $p1 = Prase-GPO ([xml](Get-GPO -Name $Gpo1 | Get-GPOReport -ReportType Xml))
  $p2 = Prase-GPO ([xml](Get-GPO -Name $Gpo2 | Get-GPOReport -ReportType Xml))

  $pp = [ordered]@{ Info = 'Name,FilterName,Domain'
      Computer = 'Name,State,Category'
      User= 'Name,State,Category'
      LinksTo = 'SOMName,SOMPath,Enabled,NoOverride'
      Account ='Name,SettingBoolean,Type'
      SecurityOptions = 'Name,SettingNumber,DName'
      UserRightsAssignment = 'Name,Members'
      RestrictedGroups = 'Name,MembeOf'
      SystemnServices = 'Name,StartupMode'
      Tasks = 'Name,Action,RunAs,LogonType,Task'
      Registry = 'Name,Key,Value'
      AuditSetting = 'PolicyTarget,SubcategoryName,SettingValue'
      DomainProfile = 'Name,Value'
      PublicProfile = 'Name,Value'
      PrivateProfile = 'Name,Value'
      RegistrySetting = 'KeyPath,AdmSetting,Value'
   }
  <#
  $pkeys = $pp.keys
  $pkeys.GetEnumerator() | % {$pp.$_ = $pp.$_ -split ','}
  
  # @($hash.GetEnumerator()) | ?{$_.key -like "*$keyword*"} | %{$hash[$_.value]=$value}
  $pp.keys | % { $pp.$_ = $pp.$_ -split ',' }

   $Global:bak | select 
   #>
  $cmp = $pp.Keys | % { "`n ---===>>> GPO : $_"
    Compare-Object ($p1.$_) ($p2.$_) -Property ($pp.$_ -split ',') | sort -Property (($pp.$_ -split ',')[0]) | ft -AutoSize }  
  "$Gpo1      vs      $Gpo2`n" + ($cmp | Out-String -Width 400) 
}

Function Prase-GPO ($GPOxml) {
 $temp = $GpoXml.GPO.Computer.ExtensionData.Extension
 [PSCustomObject]@{
  Info = $GpoXml.GPO | % { [PSCustomObject]@{ Name=$_.Name; FilterName=$_.FilterName; Domain=$_.Identifier.Domain.InnerText }}
  LinksTo = $GpoXml.GPO.LinksTo | % { [PSCustomObject]@{ SOMName=$_.SOMName; SOMPath=$_.SOMPath; Enabled=$_.Enabled; NoOverride=$_.NoOverride }}
  Computer = $temp[7].Policy | % { [PSCustomObject]@{ Name=$_.Name; State=$_.State; <#Explain=$_.Explain -replace "`n";#>  Category=$_.Category; Supported=$_.Supported; } } 
  User = $GpoXml.GPO.User.ExtensionData.Extension.Policy | % { [PSCustomObject]@{ Name = $_.Name; State=$_.State; <#Explain=$_.Explain -replace "`n";#>  Category=$_.Category; Supported=$_.Supported;  } } 
  Account = $temp[0].Account | % { [PSCustomObject]@{ Name=$_.Name; SettingBoolean=$_.SettingBoolean; Type=$_.Type } }
  UserRightsAssignment = $temp[0].UserRightsAssignment | % {  [PSCustomObject]@{ Name = $_.Name;  SID = $_.Member.SID.InnerText -join ','; Members=$_.Member.Name.InnerText -join ',';}}
  SecurityOptions = $temp[0].SecurityOptions | % {  [PSCustomObject]@{ Name=$_.KeyName; SettingNumber = $_.SettingNumber; DName=$_.Display.Name; DisplayBoolean=$_.Display.DisplayBoolean; Units=$_.Display.Units}} 
  RestrictedGroups = $temp[0].RestrictedGroups | % {  [PSCustomObject]@{ Name=$_.GroupName.Name.InnerText;  MembeOf=$_.Memberof.Name.InnerText }}
  SystemnServices = $temp[0].SystemServices | % { [PSCustomObject]@{ Name=$_.Name; StartupMode=$_.StartupMode}}
  Tasks = $temp[1].ScheduledTasks.TaskV2.Properties | % { [PSCustomObject]@{ Name=$_.Name; action=$_.Action; RunAs=$_.runAs; LogonType=$_.logonType; Task=$_.Task } }
  Registry = $temp[2].RegistrySettings.Registry.Properties | % { [PSCustomObject]@{ Name = $_.Name; action=$_.action; displayDecimal=$_.displayDecimal; default=$_.default; hive=$_.hive; key=$_.key; type=$_.type; value=$_.value; Values=$_.Values } }
  AuditSetting = $temp[3].AuditSetting | % { [PSCustomObject]@{ PolicyTarget=$_.PolicyTarget; SubcategoryName = $_.SubcategoryName; SettingValue=$_.SettingValue } }
  #RuleCollection = $temp[4].RuleCollection.type
  DomainProfile = $temp[5].DomainProfile.ChildNodes | % { [PSCustomObject]@{ Name = $_.LocalName; Value=$_.Value } }  
  PublicProfile = $temp[5].PublicProfile.ChildNodes | % { [PSCustomObject]@{ Name = $_.LocalName; Value=$_.Value } }  
  PrivateProfile =$temp[5].PrivateProfile.ChildNodes | % { [PSCustomObject]@{ Name = $_.LocalName; Value=$_.Value } } 
  RegistrySetting = $temp[7].RegistrySetting | % { [PSCustomObject]@{ KeyPath=$_.KeyPath; AdmSetting=$_.AdmSetting; Value=if (-not $_.Value) {$_.command}else{$_.Value.Name} } }
}
}

Function hl {
 param ( [string]$text, [string]$word, [System.ConsoleColor]$fc = 14, [System.ConsoleColor]$bc, [switch]$nonewline )
  #$text = ($text | Out-String).Trim()
  $s = ($text | Out-String).Trim() -split ([regex]::Escape($word))
  Write-Host $s[0] -NoNewline
  for ($i = 1; $i -lt $s.count; $i++) {  
    $params = @{ Object = $word; NoNewline = $true; ForegroundColor = $fc }
    if ($bc) { $params.BackgroundColor = $bc }
    Write-Host @params
    Write-Host $s[$i] -NoNewline
  } if (!$nonewline) { Write-Host }
}

function AskGPO($text) {
 $exist=$false
 while (-not $exist) {
   $gpo = Read-Host $text
   if ($gpo -eq $gpo1) { "Trying to use same name ???" | % {hl $_ $_ }; Write-host; continue }
   $exist = [bool](Get-Gpo -Name $gpo -ea SilentlyContinue)
   if (-not $exist) { hl "Wrong GPO name :'$gpo'" $gpo -bc Red } else { hl "GPO :'$gpo'" $gpo -fc Green };Write-host
 }
 return $gpo
}

# 22H2C-V1 Bank & Country Credit   24H2C-V1 WIN 11 B&CC

rv gpo1,gpo2 -ea SilentlyContinue
$gpo1 = AskGPO "Provide first GPO name "
$gpo2 = AskGPO "Provide second GPO name "

$path = if ($psise) { Split-Path $psise.CurrentFile.FullPath } else { $PSScriptRoot }
$date = "$(get-date -Format 'yyyy-MM-dd_HH-mm')"

Compare-GPO $gpo1 $gpo2 | tee "$path\GPO_$date.txt" | tee -Variable out

Import-Module "$path\ImportExcel\7.4.1\ImportExcel.psd1"
if (Get-Module -Name ImportExcel) { 
 $out | Export-Excel -Path "$path\GPO_$date.xlsx" ; sleep -s 1; ii "$path\GPO_$date.xlsx" 
}
