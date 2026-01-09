Describe 'Remove-OrphanedGroupMembers' {
    BeforeAll {
        . "$PSScriptRoot/Remove-OrphanedGroupMembers.ps1"
    }

    Context 'with stubbed Active Directory commands' {
        BeforeEach {
            function global:Get-ADGroup { [pscustomobject]@{ member = @('CN=Good,DC=example,DC=com','CN=Missing,DC=example,DC=com'); DistinguishedName = 'CN=Group,DC=example,DC=com'; SamAccountName = 'TestGroup' } }
            function global:Get-ADObject { param($Identity) if ($Identity -eq 'CN=Missing,DC=example,DC=com') { $null } else { [pscustomobject]@{ DistinguishedName = $Identity } } }
            $script:removedMembers = @()
            function global:Set-ADGroup { param($Identity,$Remove,[switch]$Confirm) $script:removedMembers += $Remove['member'] }
        }

        It 'removes only orphaned group members' {
            $result = Remove-OrphanedGroupMembers -GroupName 'TestGroup'

            $removedMembers | Should -Contain 'CN=Missing,DC=example,DC=com'
            $removedMembers | Should -Not -Contain 'CN=Good,DC=example,DC=com'
            $result.Removed | Should -Contain 'CN=Missing,DC=example,DC=com'
            $result.Removed | Should -Not -Contain 'CN=Good,DC=example,DC=com'
        }
    }
}
