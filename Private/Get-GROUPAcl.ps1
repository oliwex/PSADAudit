function Get-GROUPACL {
    Param(
        [Parameter(Mandatory = $true)]
        [alias("GROUP_ACL", "GroupAccessControlList")]
        [String] $groupPath
    )

    $path = "AD:\" + $groupDistinguishedName
    $acls = (Get-Acl -Path $path).Access | Select-Object ActiveDirectoryRights,AccessControlType,IdentityReference,InheritanceType,InheritanceFlags,PropagationFlags
    $info = (Get-ACL -Path $path | Select-Object Owner, Group, 'AreAccessRulesProtected', 'AreAuditRulesProtected', 'AreAccessRulesCanonical', 'AreAuditRulesCanonical')

    [PSCustomObject] @{
        'DN'                         = $groupDistinguishedName
        'Owner'                      = $info.Owner
        'Group'                      = $info.Group
        'Are Access Rules Protected' = $info.'AreAccessRulesProtected'
        'Are AuditRules Protected'   = $info.'AreAuditRulesProtected'
        'Are Access Rules Canonical' = $info.'AreAccessRulesCanonical'
        'Are Audit Rules Canonical'  = $info.'AreAuditRulesCanonical'
        'ACLs'                       = $acls
    }
}

 
