function Update-RoleEntry {
    <#
    .SYNOPSIS
    function to modify a role by removing or adding Role Entries

    .DESCRIPTION

    (If no Action is passed we assume remove)
    $roleentry should be in the form Role\Roleentry e.g. MyRole\New-DistributionGroup

    .PARAMETER RoleEntry
    Parameter description

    .PARAMETER Action
    Parameter description

    .EXAMPLE
    An example

    .NOTES
    General notes
    #>

    param(
        [Parameter()]
        $RoleEntry,

        [Parameter()]
        $Action
    )

    switch ($Action) {
        Add {
            Add-ManagementRoleEntry $RoleEntry -Confirm:$false
            break
        }
        Remove {
            Remove-ManagementRoleEntry $RoleEntry -Confirm:$false
            break
        }
        default {
            Remove-ManagementRoleEntry $RoleEntry -Confirm:$false
        }
    }
}