function Get-GPOAclSimple 
{
    Param(
        [Parameter(Mandatory = $true)]
        [Alias("GroupPolicy")]
        $groupPolicyObject
    )
    [xml]$xmlGPOReport = $groupPolicyObject.generatereport('xml')
    
    #Output
    [PsCustomObject] @{
        'Name' = $xmlGPOReport.GPO.Name
        'ACL'  = $xmlGPOReport.GPO.SecurityDescriptor.Permissions.TrusteePermissions | ForEach-Object -Process {
            [PSCustomObject] @{
                'User'            = $_.trustee.name.'#Text'
                'Permission Type' = $_.type.PermissionType
                'Inherited'       = $_.Inherited
                'Permissions'     = $_.Standard.GPOGroupedAccessEnum
            }
        }
    }
}
