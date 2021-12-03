# PSADAudit
Information about PSADAudit
[TODO:images]
# Introduction
The purpose of this project is provide information about OrganizationalUnits, Users, Groups, Computers GPO in Word document in case of organisational Audit. The module is a simple script which uses multiple functions to reach his goal. The scipt shows users in graphs, tables and images to better visualise his data.
# Technologies
* PowerShell Modules https://github.com/KevinMarquette/PSGraph
    * PSWriteWord - Author: [EvotecIT - Przemyslaw Klys](https://github.com/EvotecIT/PSWriteWord)
    * PSGraph - Author: [Kevin Marquette](https://github.com/KevinMarquette/PSGraph)
* Graphviz - simple API to create graphs in PowerShell - [Docs](https://graphviz.org/)
# Requirements
* PowerShell Min Version 5.1
* Graphviz
# Functions
## Private
### Add-Description
```
function Add-Description 
{
    Param(
        [Parameter(Mandatory = $true)]
        [Alias("DescriptionPath")]
        $pathToDescription,
        [ValidateSet("Organisational Unit", "Group", "User", "GPOPolicy", "FineGrainedPasswordPolicy","Computer")]
        [String] $descriptionType
    )

    $descriptionFileContent = Get-Content $pathToDescription | ConvertFrom-Json
    $descriptionObject=$null
    if ($descriptionType -like "Organisational Unit")
    {
        $descriptionObject=$descriptionFileContent[0].Elements.PSObject.Properties | ForEach-Object {
        "$($_.Name) - $($_.Value)"
        }
    }
    elseif ($descriptionType -like "Group") {
        $descriptionObject = $descriptionFileContent[1].Elements.PSObject.Properties | ForEach-Object {
            "$($_.Name) - $($_.Value)"
        }
    }
    elseif ($descriptionType -like "User") {
        $descriptionObject = $descriptionFileContent[2].Elements.PSObject.Properties | ForEach-Object {
            "$($_.Name) - $($_.Value)"
        }
    }
    elseif ($descriptionType -like "GPOPolicy") {
        $descriptionObject = $descriptionFileContent[3].Elements.PSObject.Properties | ForEach-Object {
            "$($_.Name) - $($_.Value)"
        }
    }
    elseif ($descriptionType -like "FineGrainedPasswordPolicy") {
        $descriptionObject = $descriptionFileContent[4].Elements.PSObject.Properties | ForEach-Object {
            "$($_.Name) - $($_.Value)"
        }
    }
    elseif ($descriptionType -like "Computer") {
        $descriptionObject = $descriptionFileContent[5].Elements.PSObject.Properties | ForEach-Object {
            "$($_.Name) - $($_.Value)"
        }
    }
    Add-WordList -WordDocument $reportFile -ListType Bulleted -ListData $descriptionObject -Supress $true -Verbose

}
```
### Add-Description
```
function Add-WordChart {
    Param(
        [Parameter(Mandatory = $true)]
        [alias("CType")]
        [ValidateSet("Piechart", "Barchart")]
        [String] $chartType,
        [Parameter(Mandatory = $true)]
        [alias("CData")]
        $chartData,
        [Parameter(Mandatory = $true)]
        [alias("STitle")]
        [String] $sectionTitle,
        [Parameter(Mandatory = $true)]
        [alias("CTitle")]
        [String] $chartTitle

    )
    Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text $sectionTitle -Supress $true
    [array] $Names = foreach ($nameTMP in $chartData) {
        "$($nameTMP.Name) - [$($nameTMP.Values)]"
    }
    if ($chartType -like "*PieChart*") {    
        Add-WordPieChart -WordDocument $reportFile -ChartName $chartTitle -ChartLegendPosition Bottom -ChartLegendOverlay $false -Names $Names -Values $([array]$chartData.Values)
    }
    else {
        Add-WordBarChart -WordDocument $reportFile -ChartName $chartTitle -ChartLegendPosition Bottom -ChartLegendOverlay $false -Names $Names -Values $([array]$chartData.Values) -BarDirection Column   
    }
}
```
### Get-ComputerAcl
```
function Get-COMPUTERAcl {
    Param(
        [Parameter(Mandatory = $true)]
        [alias("COMPUTER_ACL", "UserAccessControlList")]
        [String] $computerPath
    )

    $path = "AD:\" + $computerPath
    $acls = (Get-Acl -Path $path).Access | Select-Object ActiveDirectoryRights, AccessControlType, IdentityReference, InheritanceType, InheritanceFlags, PropagationFlags
    $info = (Get-ACL -Path $path | Select-Object Owner, Group, 'AreAccessRulesProtected', 'AreAuditRulesProtected', 'AreAccessRulesCanonical', 'AreAuditRulesCanonical')

    [PSCustomObject] @{
        'DN'                         = $computerPath
        'Owner'                      = $info.Owner
        'Group'                      = $info.Group
        'Are Access Rules Protected' = $info.'AreAccessRulesProtected'
        'Are AuditRules Protected'   = $info.'AreAuditRulesProtected'
        'Are Access Rules Canonical' = $info.'AreAccessRulesCanonical'
        'Are Audit Rules Canonical'  = $info.'AreAuditRulesCanonical'
        'ACLs'                       = $acls
    }
}
```
### Get-ComputerInformation
```
function Get-COMPUTERInformation {
    $computerData = Get-ADComputer -Filter * -Properties *
    $computerOutput = foreach ($data in $computerData) {
        #AccountExpires,
        [PSCustomObject] @{
            'AccountExpirationDate'                = $data.AccountExpirationDate
            'AccountLockoutTime'                   = $data.AccountLockoutTime
            'AccountNotDelegated'                  = $data.AccountNotDelegated
            'AllowReversiblePasswordEncryption'    = $data.AllowReversiblePasswordEncryption
            'AuthenticationPolicy'                 = $data.AuthenticationPolicy
            'AuthenticationPolicySilo'             = $data.AuthenticationPolicySilo
            'BadLogonCount'                        = $data.BadLogonCount
            'CannotChangePassword'                 = $data.CannotChangePassword
            'CanonicalName'                        = $data.CanonicalName
            'Certificates'                         = $data.Certificates
            'CommonName'                           = $data.CommonName
            'CodePage'                             = $data.codepage
            'CountryCode'                          = $data.CountryCode
            'Description'                          = $data.Description
            'DisplayName'                          = $data.DisplayName
            'DistinguishedName'                    = $data.DistinguishedName
            'DNSHostName'                          = $data.DNSHostName
            'DoesNotRequirePreAuth'                = $data.DoesNotRequirePreAuth
            'Enabled'                              = $data.Enabled
            'HomeDirRequired'                      = $data.HomeDirRequired
            'HomePage'                             = $data.HomePage
            'InstanceType'                         = $data.instanceType
            'IP4'                                  = $data.IPv4Address
            'IP6'                                  = $data.IPv6Address
            'IsCriticalSystemObject'               = $data.isCriticalSystemObject
            'KerberosEncryptionType'               = $data."msDS-SupportedEncryptionTypes"
            'LastBadPasswordAttempt'               = $data.LastBadPasswordAttempt
            'LastKnownParent'                      = $data.LastKnownParent
            'LastLogonDate'                        = $data.LastLogonDate
            'LocalPolicyFlags'                     = $data.LocalPolicyFlags
            'Location'                             = $data.Location
            'LockedOut'                            = $data.LockedOut
            'LogonCount'                           = $data.LogonCount
            'ManagedBy'                            = $data.ManagedBy
            'MemberOf'                             = $data.MemberOf
            'Name'                                 = $data.Name
            'ObjectCategory'                       = $data.ObjectCategory
            'ObjectClass'                          = $data.ObjectClass
            'ObjectGUID'                           = $data.ObjectGUID
            'OperatingSystem'                      = $data.OperatingSystem
            'OperatingSystemHotfix'                = $data.OperatingSystemHotfix
            'OperatingSystemServicePack'           = $data.OperatinSystemServicePack
            'OperatingSystemVersion'               = $data.OperatingSystemVersion
            'PasswordExpired'                      = $data.PasswordExpired
            'PasswordLastSet'                      = $data.PasswordLastSet
            'PasswordNeverExpires'                 = $data.PasswordNeverExpires
            'PasswordNotRequired'                  = $data.PasswordNotRequired
            'PrimaryGroup'                         = $data.PrimaryGroup
            'PrincipalsAllowedToDelegateToAccount' = $data.PrincipalsAllowedToDelegateToAccount
            'ProtectedFromAccidentalDeletion'      = $data.ProtectedFromAccidentalDeletion
            'SamAccountName'                       = $data.SamAccountName
            'SamAccountType'                       = $data.SamAccountType
            'ServiceAccount'                       = $data.ServiceAccount
            'ServicePrincipalName'                 = $data.ServicePrincipalName
            'ServicePrincipalNames'                = $data.ServicePrincipalNames
            'SID'                                  = $data.SID
            'SIDHistory'                           = $data.SIDHistory
            'TrustedForDelegation'                 = $data.TrustedForDelegation
            'TrustedToAuthForDelegation'           = $data.TrustedToAuthForDelegation
            'UseDESKeyOnly'                        = $data.UseDESKeyOnly
            'userAccountControl'                   = $data.userAccountControl
            'UserCertificate'                      = $data.UserCertificate
            'UserPrincipalName'                    = $data.UserPrincipalName
            'USNChanged'                           = $data.uSNChanged
            'USNCreated'                           = $data.uSNCreated
            'WhenChanged'                          = $data.whenChanged
            'WhenCreated'                          = $data.whenCreated
        }
    }
    $computerOutput 
}
```
### Get-FineGrainedPolicies
```
function Get-FineGrainedPolicies {
    $fineGrainedPoliciesData = Get-ADFineGrainedPasswordPolicy -Filter * -Server $($Env:USERDNSDOMAIN)
    $fineGrainedPolicies = foreach ($policy in $fineGrainedPoliciesData) {
        [PsCustomObject] @{
            'Name'                          = $policy.Name
            'Complexity Enabled'            = $policy.ComplexityEnabled
            'Lockout Duration'              = $policy.LockoutDuration
            'Lockout Observation Window'    = $policy.LockoutObservationWindow
            'Lockout Threshold'             = $policy.LockoutThreshold
            'Max Password Age'              = $policy.MaxPasswordAge
            'Min Password Length'           = $policy.MinPasswordLength
            'Min Password Age'              = $policy.MinPasswordAge
            'Password History Count'        = $policy.PasswordHistoryCount
            'Reversible Encryption Enabled' = $policy.ReversibleEncryptionEnabled
            'Precedence'                    = $policy.Precedence
            'Applies To'                    = $policy.AppliesTo 
            'Distinguished Name'            = $policy.DistinguishedName
        }
    }
    return $fineGrainedPolicies
}
```
### Get-GPOAclExtended
```
function Get-GPOAclExtended
{
    Param(
        [Parameter(Mandatory = $true)]
        [alias("GPO_ACL", "GPOAccessControlList")]
        [String] $gpoPath
    )

    $path = "AD:\" + $gpoPath
    $acls = (Get-Acl -Path $path).Access | Select-Object ActiveDirectoryRights, AccessControlType, IdentityReference, InheritanceType, InheritanceFlags, PropagationFlags
    $info = (Get-ACL -Path $path | Select-Object Owner, Group, 'AreAccessRulesProtected', 'AreAuditRulesProtected', 'AreAccessRulesCanonical', 'AreAuditRulesCanonical')

    [PSCustomObject] @{
        'DN'                         = $gpoPath
        'Owner'                      = $info.Owner
        'Group'                      = $info.Group
        'Are Access Rules Protected' = $info.'AreAccessRulesProtected'
        'Are AuditRules Protected'   = $info.'AreAuditRulesProtected'
        'Are Access Rules Canonical' = $info.'AreAccessRulesCanonical'
        'Are Audit Rules Canonical'  = $info.'AreAuditRulesCanonical'
        'ACLs'                       = $acls
    }
}
```
### Get-GPOAclSimple
```
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
```
### Get-GPOPolicy
```
function Get-GPOPolicy {
    Param(
        [Parameter(Mandatory = $true)]
        [Alias("GPO_Object")]
        $groupPolicyObject
    )
   
    [xml]$xmlGPOReport = $groupPolicyObject.generatereport('xml')
    #GPO version
    if (($xmlGPOReport.GPO.Computer.VersionDirectory -eq 0) -and ($xmlGPOReport.GPO.Computer.VersionSysvol -eq 0)) {
        $computerSettings = "NeverModified"
    } 
    else {
        $computerSettings = "Modified"
    }
    if (($xmlGPOReport.GPO.User.VersionDirectory -eq 0) -and ($xmlGPOReport.GPO.User.VersionSysvol -eq 0)) {
        $userSettings = "NeverModified"
    } 
    else {
        $userSettings = "Modified"
    }

    #GPO content
    if ($null -eq $xmlGPOReport.GPO.User.ExtensionData) {
        $userSettingsConfigured = $false
    } 
    else {
        $userSettingsConfigured = $true
    }
    if ($null -eq $xmlGPOReport.GPO.Computer.ExtensionData) {
        $computerSettingsConfigured = $false
    } 
    else {
        $computerSettingsConfigured = $true
    }
    #Output
    [PsCustomObject] @{
        'Name'                 = $xmlGPOReport.GPO.Name
        'Links'                = $xmlGPOReport.GPO.LinksTo | Select-Object -ExpandProperty SOMPath
        'HasComputerSettings'  = $computerSettingsConfigured
        'HasUserSettings'      = $userSettingsConfigured
        'UserEnabled'          = $xmlGPOReport.GPO.User.Enabled
        'ComputerEnabled'      = $xmlGPOReport.GPO.Computer.Enabled
        'ComputerSettings'     = $computerSettings
        'UserSettings'         = $userSettings
        'GpoStatus'            = $groupPolicyObject.GpoStatus
        'CreationTime'         = $groupPolicyObject.CreationTime
        'ModificationTime'     = $groupPolicyObject.ModificationTime
        'WMIFilter'            = $groupPolicyObject.WmiFilter.name
        'WMIFilterDescription' = $groupPolicyObject.WmiFilter.Description
        'Path'                 = $groupPolicyObject.Path
        'GUID'                 = $groupPolicyObject.Id
    }

}
```
### Get-GraphImage
```
function Get-GraphImage {
    Param(
        [Alias("GraphRoot")]
        $root,
        [Alias("GraphMiddle")]
        $middle, 
        [Alias("GraphLeaf")]
        $leaf,
        [Alias("BasePathToGraphImage")]
        $pathToImage
    )

    $imagePath = Join-Path -Path $pathToImage -ChildPath "$middle.png"
    $graphTMP=$null
    if ($null -eq $root) #not have boss
    {
        $graphTMP = graph g {
            edge -From $middle -To $leaf
        }    
    }
    elseif ($null -eq $leaf) #not have employees below
    {
        $graphTMP = graph g {
            edge -From $root -To $middle
        } 
    }
    else #have boss and employees
    {
        $graphTMP = graph g {
                    edge -From $root -To $middle
                    edge -From $middle -To $leaf
                }
    }
    
    $vizPath = Join-Path -Path $pathToImage -ChildPath "$middle.vz"
    Set-Content -Path $vizPath -Value $graphTMP
    Export-PSGraph -Source $vizPath -Destination $imagePath

    #cleaning
    Remove-Item -Path $vizPath

    $imagePath
}
```
### Get-GroupAcl
```
function Get-GROUPAcl {
    Param(
        [Parameter(Mandatory = $true)]
        [alias("GROUP_ACL", "GroupAccessControlList")]
        [String] $groupPath
    )

    $path = "AD:\" + $groupPath
    $acls = (Get-Acl -Path $path).Access | Select-Object ActiveDirectoryRights,AccessControlType,IdentityReference,InheritanceType,InheritanceFlags,PropagationFlags
    $info = (Get-ACL -Path $path | Select-Object Owner, Group, 'AreAccessRulesProtected', 'AreAuditRulesProtected', 'AreAccessRulesCanonical', 'AreAuditRulesCanonical')

    [PSCustomObject] @{
        'DN'                         = $groupPath
        'Owner'                      = $info.Owner
        'Group'                      = $info.Group
        'Are Access Rules Protected' = $info.'AreAccessRulesProtected'
        'Are AuditRules Protected'   = $info.'AreAuditRulesProtected'
        'Are Access Rules Canonical' = $info.'AreAccessRulesCanonical'
        'Are Audit Rules Canonical'  = $info.'AreAuditRulesCanonical'
        'ACLs'                       = $acls
    }
}

```
### Get-GroupInformation
```
function Get-GROUPInformation {
    $groupData = Get-ADGroup -Filter * -Properties *
    $groupOutput = foreach ($data in $groupData) {
        [PSCustomObject] @{
            'CanonicalName'                   = $data.CanonicalName
            'Common Name'                     = $data.cn
            'Description'                     = $data.Description
            'DisplayName'                     = $data.DisplayName
            'DistinguishedName'               = $data.DistinguishedName
            'GroupCategory'                   = $data.GroupCategory
            'GroupScope'                      = $data.GroupScope
            'GroupType'                       = $data.groupType
            'HomePage'                        = $data.HomePage
            'InstanceType'                    = $data.instanceType
            'ManagedBy'                       = $data.ManagedBy
            'MemberOf'                        = $data.MemberOf
            'Members'                         = $data.Members
            'Name'                            = $data.Name
            'ObjectCategory'                  = $data.ObjectCategory
            'ObjectClass'                     = $data.ObjectClass
            'ObjectGuid'                      = $data.ObjectGuid
            'ProtectedFromAccidentalDeletion' = $data.ProtectedFromAccidentalDeletion
            'SamAccountName'                  = $data.SamAccountName
            'SAMAccountType'                  = $data.sAMAccountType
            'SID'                             = $data.SID
            'SIDHistory'                      = $data.SIDHistory
            'USNChanged'                      = $data.uSNChanged
            'USNCreated'                      = $data.uSNCreated
            'WhenChanged'                     = $data.whenChanged
            'WhenCreated'                     = $data.whenCreated
        }
    }
    $groupOutput
}
```
### Get-OUAcl
```
function Get-OUAcl {
    Param(
        [Parameter(Mandatory = $true)]
        [alias("OU_ACL", "OrganisationalUnitAccessControlList")]
        [String] $ouPath
    )
    $path = "AD:\" + $ouPath
    $acls = (Get-Acl -Path $path).Access | Select-Object ActiveDirectoryRights,AccessControlType,IdentityReference,InheritanceType,InheritanceFlags,PropagationFlags

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
```
### Get-OUInformation
```
function Get-OUInformation {
    $ouData = Get-ADOrganizationalUnit -Filter * -Properties * 

    $ouOutput = foreach ($data in $ouData) {
        [PSCustomObject] @{
            'CanonicalName'                   = $data.CanonicalName
            'City'                            = $data.City
            'Common Name'                     = $data.cn
            'Country'                         = $data.Country
            'Description'                     = $data.Description
            'DisplayName'                     = $data.DisplayName
            'DistinguishedName'               = $data.DistinguishedName
            'GPLink'                          = $data.gPLink
            'InstanceType'                    = $data.instanceType
            'IsCriticalSystemObject'          = $data.isCriticalSystemObject
            'LastKnownParent'                 = $data.LastKnownParent
            'LinkedGroupPolicyObjects'        = $data.LinkedGroupPolicyObjects
            'ManagedBy'                       = $data.ManagedBy
            'Modified'                        = $data.Modified
            'Name'                            = $data.Name
            'ObjectCategory'                  = $data.ObjectCategory
            'ObjectClass'                     = $data.ObjectClass
            'ObjectGuid'                      = $data.ObjectGuid
            'PostalCode'                      = $data.PostalCode
            'ProtectedFromAccidentalDeletion' = $data.ProtectedFromAccidentalDeletion
            'ShowInAdvancedViewOnly'          = $data.showInAdvancedViewOnly
            'State'                           = $data.State
            'StreetAddress'                   = $data.StreetAddress
            'USNChanged'                      = $data.uSNChanged
            'USNCreated'                      = $data.uSNCreated
            'WhenChanged'                     = $data.whenChanged
            'WhenCreated'                     = $data.whenCreated
        }
    }
    $ouOutput
}
```
### Get-ReportFolders
```
function Get-ReportFolders {
    Param(
        [Parameter(Mandatory = $true)]
        [Alias("BasePath")]
        [string]$reportPath,
        [Alias("GraphFoldersHashtable")]
        $graphFolders
    )
    $graphFoldersOutput=@{}
    $($graphFolders.Keys) | ForEach-Object {
        $folderPath = Join-Path -Path $reportPath -ChildPath $graphFolders[$_]
        New-Item -Path $folderPath -ItemType Directory | Out-Null
        $graphFoldersOutput.Add($_, $folderPath)
    }
    $graphFoldersOutput
}
```
### Get-UserAcl
```
function Get-USERAcl 
{
    Param(
        [Parameter(Mandatory = $true)]
        [alias("USER_ACL", "UserAccessControlList")]
        [String] $userPath
    )

    $path = "AD:\" + $userPath
    $acls = (Get-Acl -Path $path).Access | Select-Object ActiveDirectoryRights, AccessControlType, IdentityReference, InheritanceType, InheritanceFlags, PropagationFlags
    $info = (Get-ACL -Path $path | Select-Object Owner, Group, 'AreAccessRulesProtected', 'AreAuditRulesProtected', 'AreAccessRulesCanonical', 'AreAuditRulesCanonical')

    [PSCustomObject] @{
        'DN'                         = $userPath
        'Owner'                      = $info.Owner
        'Group'                      = $info.Group
        'Are Access Rules Protected' = $info.'AreAccessRulesProtected'
        'Are AuditRules Protected'   = $info.'AreAuditRulesProtected'
        'Are Access Rules Canonical' = $info.'AreAccessRulesCanonical'
        'Are Audit Rules Canonical'  = $info.'AreAuditRulesCanonical'
        'ACLs'                       = $acls
    }
}

```
### Get-UserInformation
```
function Get-USERInformation {
    $userData = Get-ADUser -Filter * -Properties *
    $userOutput = foreach ($data in $userData) {
        [PSCustomObject] @{
            'AccountExpirationDate'             = $data.AccountExpirationDate
            'AccountLockoutTime'                = $data.AccountLockoutTime
            'AccountNotDelegated'               = $data.AccountNotDelegated
            'AllowReversiblePasswordEncryption' = $data.AllowReversiblePasswordEncryption
            'BadLogonCount'                     = $data.BadLogonCount
            'CannotChangePassword'              = $data.CannotChangePassword
            'CanonicalName'                     = $data.CanonicalName
            'Certificates'                      = $data.Certificates
            'ChangePasswordAtLogon'             = $data.ChangePasswordAtLogon
            'City'                              = $data.City
            'CommonName'                        = $data.cn
            'Company'                           = $data.Company
            'Country'                           = $data.Country
            'DesktopProfile'                    = $data.DesktopProfile
            'Department'                        = $data.Department
            'Description'                       = $data.Description
            'DirectReports'                     = $data.DirectReports
            'DisplayName'                       = $data.DisplayName
            'DistinguishedName'                 = $data.DistinguishedName
            'Division'                          = $data.Division
            'DoesNotRequirePreAuth'             = $data.DoesNotRequirePreAuth
            'EmailAddress'                      = $data.EmailAddress
            'EmployeeID'                        = $data.EmployeeID
            'EmployeeNumber'                    = $data.EmployeeNumber
            'Enabled'                           = $data.Enabled
            'Fax'                               = $data.Fax
            'GivenName'                         = $data.GivenName
            'GroupMembershipSAM'                = $data.groupMembershipSAM
            'HomeDirectory'                     = $data.HomeDirectory
            'HomeDirRequired'                   = $data.HomeDirEnabled
            'HomeDrive'                         = $data.HomeDrive
            'HomePage'                          = $data.HomePage
            'HomePhone'                         = $data.HomePhone
            'LastBadPasswordAttempt'            = $data.LastBadPasswordAttempt
            'LastKnownParent'                   = $data.LastKnownParent
            'LastLogOn'                         = $data.LastLogOn
            'LastLogOff'                        = $data.LastLogOff
            'LastLogonDate'                     = $data.LastLogonDate
            'LockedOut'                         = $data.LockedOut
            'LockoutTime'                       = $data.LockoutTime
            'LogonHours'                        = $data.LogonHours
            'LogonWorkstations'                 = $data.LogonWorkstations
            'Manager'                           = $data.Manager
            'MemberOf'                          = $data.MemberOf
            'MobilePhone'                       = $data.MobilePhone
            'Name'                              = $data.Name
            'ObjectCategory'                    = $data.ObjectCategory
            'ObjectClass'                       = $data.ObjectClass
            'ObjectGuid'                        = $data.ObjectGuid
            'Office'                            = $data.Office
            'OfficePhone'                       = $data.OfficePhone
            'Organization'                      = $data.Organization
            'OtherName'                         = $data.OtherName
            'PasswordExpired'                   = $data.PasswordExpired
            'PasswordLastSet'                   = $data.PasswordLastSet
            'PasswordNeverExpires'              = $data.PasswordNeverExpires
            'PasswordNotRequired'               = $data.PasswordNotRequired
            'POBox'                             = $data.POBox
            'PostalCode'                        = $data.PostalCode
            'PrimaryGroup'                      = $data.PrimaryGroup
            'ProfilePath'                       = $data.ProfilePath
            'ProtectedFromAccidentalDeletion'   = $data.ProtectedFromAccidentalDeletion
            'SamAccountName'                    = $data.SamAccountName
            'ScriptPath'                        = $data.ScriptPath
            'ShowInAdvancedViewOnly'            = $data.showInAdvancedViewOnly
            'ServicePrincipalName'              = $data.ServicePrincipalName
            'SID'                               = $data.SID 
            'SIDHistory'                        = $data.SIDHistory
            'SmartcardLogonRequired'            = $data.SmartcardLogonRequired
            'State'                             = $data.State 
            'StreetAddress'                     = $data.StreetAddress
            'Surname'                           = $data.Surname 
            'ThumbnailPhoto'                    = $data.ThumbnailPhoto
            'ThumbnailLogo'                     = $data.ThumbnailLogo
            'Title'                             = $data.Title
            'TrustedForDelegation'              = $data.TrustedForDelegation
            'TrustedToAuthForDelegation'        = $data.TrustedToAuthForDelegation
            'UserAccountControl'                = $data.UserAccountControl
            'UseDESKeyOnly'                     = $data.UseDESKeyOnly
            'UserPrincipalName'                 = $data.UserPrincipalName
            'whenCreated'                       = $data.whenCreated
            'whenChanged'                       = $data.whenChanged
        }
    }
    $userOutput
}
```
### Install-Chocolatey
```
function Install-Chocolatey
{
    #Install-Chocolatey
    Set-ExecutionPolicy Bypass -Scope Process -Force; 
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; 
    Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
    #Install graphviz
    choco install graphviz
}
```
### New-Workplace
```
function New-Workplace
{
    do
    {
        $workplacePath=Read-Host "Get Path for Report:"
        $isPathCorrect=(([string]::IsNullOrWhiteSpace($workplacePath)) -or (Test-Path -Path $workplacePath -PathType Container))
    }
    while ($isPathCorrect)

    New-Item -Path $workplacePath -ItemType Directory | Out-Null
    $workplacePath
}
```
## Public
# Results
[Example results]
# Examples
[Example gif or movie]