Describe 'Remove-OrphanedGroupMembers' {
    BeforeAll {
        . "$PSScriptRoot/Remove-OrphanedGroupMembers.ps1"
    }

    Context 'with stubbed Active Directory commands' {
        BeforeEach {
            function global:Get-ADGroup {
                param($Identity, [string[]]$Properties)
                $Identity | Should -Be 'TestGroup'
                $Properties | Should -Contain 'member'
                [pscustomobject]@{
                    member = @('CN=Good,DC=example,DC=com','CN=Missing,DC=example,DC=com')
                    DistinguishedName = 'CN=Group,DC=example,DC=com'
                    SamAccountName = 'TestGroup'
                }
            }
            function global:Get-ADObject {
                param($Identity)
                $Identity | Should -Not -BeNullOrEmpty
                if ($Identity -eq 'CN=Missing,DC=example,DC=com') {
                    $null
                } else {
                    [pscustomobject]@{ DistinguishedName = $Identity }
                }
            }
            $script:removedMembers = [System.Collections.ArrayList]@()
            function global:Set-ADGroup {
                param($Identity, $Remove, [switch]$Confirm)
                $Identity | Should -Be 'CN=Group,DC=example,DC=com'
                $Remove['member'] | Should -Not -BeNullOrEmpty
                $null = $script:removedMembers.Add($Remove['member'])
            }
        }

        It 'removes only orphaned group members' {
            $result = Remove-OrphanedGroupMembers -GroupName 'TestGroup'

            $removedMembers | Should -Contain 'CN=Missing,DC=example,DC=com'
            $removedMembers | Should -Not -Contain 'CN=Good,DC=example,DC=com'
            $result.Orphaned | Should -Contain 'CN=Missing,DC=example,DC=com'
            $result.Orphaned | Should -Not -Contain 'CN=Good,DC=example,DC=com'
            $result.Removed | Should -Contain 'CN=Missing,DC=example,DC=com'
            $result.Removed | Should -Not -Contain 'CN=Good,DC=example,DC=com'
        }
    }
}
