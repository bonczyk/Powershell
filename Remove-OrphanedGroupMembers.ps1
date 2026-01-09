function Remove-OrphanedGroupMembers {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupName
    )

    $group = Get-ADGroup -Identity $GroupName -Properties member -ErrorAction Stop

    $orphanedMembers = foreach ($dn in $group.member) {
        if (-not (Get-ADObject -Identity $dn -ErrorAction SilentlyContinue)) {
            $dn
        }
    }

    foreach ($dn in $orphanedMembers) {
        if ($PSCmdlet.ShouldProcess($GroupName, "Remove member $dn")) {
            Set-ADGroup -Identity $group.DistinguishedName -Remove @{ member = $dn } -Confirm:$false
        }
    }

    [PSCustomObject]@{
        Group   = $group.SamAccountName
        Removed = $orphanedMembers
    }
}
