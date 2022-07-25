function Set-WorkWeekSchedule($ProgramName,$CollectionName,$time) {

#Get-CMPackageDeployment -ProgramName $ProgramName | Select-Object PackageID -ExpandProperty AssignedSchedule

 $a = 1..5 | % { New-CMSchedule -DayOfWeek $_ -Start (Get-Date -F "dd/MM/yy $time") }

Get-CMDeployment -ProgramName $ProgramName -CollectionName $CollectionName | Set-CMPackageDeployment -StandardProgramName $ProgramName -Schedule $a 

}

 

Set-WorkWeekSchedule -ProgramName wol2 -CollectionName "WOL backup test" -time 06:32
