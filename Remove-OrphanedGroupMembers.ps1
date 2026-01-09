function Remove-OrphanedGroupMembers {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupName
    )

    try {
        $group = Get-ADGroup -Identity $GroupName -Properties member -ErrorAction Stop
    } catch {
        throw "Failed to retrieve group '$GroupName': $($_.Exception.Message)"
    }

    $orphanedMembers = @(
        foreach ($dn in $group.member) {
            if (-not (Get-ADObject -Identity $dn -ErrorAction SilentlyContinue)) {
                $dn
            }
        }
    )

    $removedMembers = [System.Collections.ArrayList]@()

    foreach ($dn in $orphanedMembers) {
        if ($PSCmdlet.ShouldProcess($GroupName, "Remove member $dn")) {
            Set-ADGroup -Identity $group.DistinguishedName -Remove @{ member = $dn }
            $null = $removedMembers.Add($dn)
        }
    }

    [PSCustomObject]@{
        Group   = $group.SamAccountName
        Orphaned = $orphanedMembers
        Removed  = @($removedMembers)
    }
}
