function Get-OUACL {
    Param(
        [Parameter(Mandatory = $true)]
        [alias("OU_ACL", "OrganisationalUnitAccessControlList")]
        [String] $ouPath
    )
    $path = "AD:\" + $ouPath
    $acls = (Get-Acl -Path $path).Access | Where-Object { $_.IsInherited -eq $false }

    $info = (Get-ACL -Path $path | Select-Object Owner, Group, 'AreAccessRulesProtected', 'AreAuditRulesProtected', 'AreAccessRulesCanonical', 'AreAuditRulesCanonical')
    
    [PSCustomObject] @{
        'DN'                         = $ouPath
        'Owner'                      = $info.Owner
        'Group'                      = $info.Group
        'Are Access Rules Protected' = $info.'AreAccessRulesProtected'
        'Are AuditRules Protected'   = $info.'AreAuditRulesProtected'
        'Are Access Rules Canonical' = $info.'AreAccessRulesCanonical'
        'Are Audit Rules Canonical'  = $info.'AreAuditRulesCanonical'
        'ACLs'                       = $acls
    }
}