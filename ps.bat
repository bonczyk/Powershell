@(set "0=%~f0"^)#) & powershell -nop -c iex([io.file]::ReadAllText($env:0)) & exit/b

Write-host 'same window'; pause

$_Paste_in_Powershell = { Write-host 'new window'; pause }
start powershell -args "-nop -c & {`n`n$($_Paste_in_Powershell-replace'"','\"')}" #-verb runas
