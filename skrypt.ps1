
#TODO:Create unified function for table
#TODO:Create unified function for list
##########################################################################################
#                                GLOBAL VARIABLES                                        #
##########################################################################################
$basePath = "C:\reporty\"
$graphFolders = @{
    GPO   = "GPO_Graph\"
    OU    = "OU_Graph\"
    FGPP  = "FGPP_Graph\"
    GROUP = "GROUP_Graph\"
    USERS = "USERS_Graph\"
}
##########################################################################################
function Get-OUInformation 
{
    $ouData = Get-ADOrganizationalUnit -Filter * -Properties * 

    $ouOutput = foreach ($data in $ouData) 
    {
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
function Get-GROUPInformation 
{
    $groupData = Get-ADGroup -Filter * -Properties *
    $groupOutput = foreach ($data in $groupData) 
    {
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
function Get-USERInformation 
{
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
            'DirectReport'                      = $data.DirectReports
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

#TEST
function Get-ComputerInformation 
{
    $computerData = Get-ADComputer -Filter * -Properties *
    $computerOutput = foreach ($data in $userData) 
    {
    #AccountExpires,
    [PSCustomObject] @{
        'AccountExpirationDate' = $data.AccountExpirationDate
        'AccountLockoutTime'    = $data.AccountLockoutTime
        'AccountNotDelegated'   = $data.AccountNotDelegated
        'AllowReversiblePasswordEncryption' = $data.AllowReversiblePasswordEncryption
        'AuthenticationPolicy'   = $data.AuthenticationPolicy
        'AuthenticationPolicySilo'  = $data.AuthenticationPolicySilo
        'BadLogonCount' = $data.BadLogonCount
        'CannotChangePassword'  = $data.CannotChangePassword
        'CanonicalName' = $data.CanonicalName
        'Certificates' = $data.Certificates
        'CommonName' = $data.CommonName
        'CodePage' = $data.codepage
            'CountryCode' = $data.CountryCode
        'Description' = $data.Description
        'DisplayName' = $data.DisplayName
        'DistinguishedName' = $data.DistinguishedName
        'DNSHostName' = $data.DNSHostName
        'DoesNotRequirePreAuth' = $data.DoesNotRequirePreAuth
        'Enabled' = $data.Enabled
        'HomeDirRequired' = $data.HomeDirRequired
        'HomePage' = $data.HomePage
        'InstanceType' = $data.instanceType
        'IP4' = $data.IPv4Address
        'IP6' = $data.IPv6Address
        'IsCriticalSystemObject' = $data.isCriticalSystemObject
        'KerberosEncryptionType' = $data.KerberosEncryptionType
        'LastBadPasswordAttempt' = $data.LastBadPasswordAttempt
        'LastKnownParent' = $data.LastKnownParent
        'LastLogonDate' = $data.LastLogonDate
            'LocalPolicyFlags' = $data.LocalPolicyFlags
        'Location' = $data.Location
        'LockedOut' = $data.LockedOut
        'LogonCount' = $data.LogonCount
        'ManagedBy' = $data.ManagedBy
        'MemberOf' = $data.MemberOf
        'Name' = $data.Name
        'ObjectCategory' = $data.ObjectCategory
        'ObjectClass' = $data.ObjectClass
        'ObjectGUID' = $data.ObjectGUID
        'OperatingSystem' = $data.OperatingSystem
        'OperatingSystemHotfix' = $data.OperatingSystemHotfix
        'OperatingSystemServicePack' = $data.OperatinSystemServicePack
        'OperatingSystemVersion' = $data.OperatingSystemVersion
        'PasswordExpired' = $data.PasswordExpired
        'PasswordLastSet' = $data.PasswordLastSet
        'PasswordNeverExpires' = $data.PasswordNeverExpires
        'PasswordNotRequired' = $data.PasswordNotRequired
        'PrimaryGroup' = $data.PrimaryGroup
        'PrincipalsAllowedToDelegateToAccount' = $data.PrincipalsAllowedToDelegateToAccount
        'ProtectedFromAccidentalDeletion' = $data.ProtectedFromAccidentalDeletion
        'SamAccountName' = $data.SamAccountName
        'SamAccountType' = $data.SamAccountType
        'ServiceAccount' = $data.ServiceAccount
        'ServicePrincipalName' = $data.ServicePrincipalName
        'ServicePrincipalNames' = $data.ServicePrincipalNames
        'SID' = $data.SID
        'SIDHistory' = $data.SIDHistory
        'TrustedForDelegation' = $data.TrustedForDelegation
        'TrustedToAuthForDelegation' = $data.TrustedToAuthForDelegation
        'UseDESKeyOnly' = $data.UseDESKeyOnly
        'userAccountControl' = $data.userAccountControl
        'UserCertificate' = $data.UserCertificate
        'UserPrincipalName' = $data.UserPrincipalName
        'USNChanged'    = $data.uSNChanged
        'USNCreated'    = $data.uSNCreated
        'WhenChanged'   = $data.whenChanged
        'WhenCreated'   = $data.whenCreated
        }
    }
$computerOutput 
}

function Get-GPOPolicy 
{

   $groupPolicyObjects =  Get-GPO -Domain $($Env:USERDNSDOMAIN) -All 
   $groupPolicyOutput=foreach ($groupPolicyObject in $groupPolicyObjects)
   {

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
    $groupPolicyOutput
}
function Get-GPOAcl {
    Param(
        [Parameter(Mandatory = $true)]
        [alias("GPOObject", "GroupPolicyObject")]
        $groupPolicyObjectAcl
    )

    [xml]$xmlGPOReport = $groupPolicyObjectAcl.generatereport('xml')

    #Output
    [PsCustomObject] @{
        'Name' = $xmlGPOReport.GPO.Name
        'ACLs' = $xmlGPOReport.GPO.SecurityDescriptor.Permissions.TrusteePermissions | ForEach-Object -Process {
            New-Object -TypeName PSObject -Property @{
                'User'            = $_.trustee.name.'#Text'
                'Permission Type' = $_.type.PermissionType
                'Inherited'       = $_.Inherited
                'Permissions'     = $_.Standard.GPOGroupedAccessEnum
            }
        }
    }
}
function Get-GraphImage {
    Param(
        [Parameter(Mandatory = $true)]
        [Alias("GraphRoot")]
        $root, 
        [Alias("GraphLeaf")]
        $leaf, 
        [Alias("BasePathToGraphImage")]
        $pathToImage
    )

    $imagePath = Join-Path -Path $pathToImage -ChildPath "$root.png"
        
    $graphTMP = graph g {
        edge -from $root -To $leaf
    }
    
    $vizPath = Join-Path -Path $pathToImage -ChildPath "$root.vz"
    Set-Content -Path $vizPath -Value $graphTMP
    Export-PSGraph -Source $vizPath -Destination $imagePath

    #cleaning
    Remove-Item -Path $vizPath

    $imagePath
}
#TODO:Analysis
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
##########################################################################################
#                                TOOL FUNCTIONS                                          #
##########################################################################################
function Get-ReportFolders {
    Param(
        [Parameter(Mandatory = $true)]
        [Alias("BasePath")]
        [string]$reportPath,
        [Alias("GraphFoldersHashtable")]
        $graphFolders
    )

    foreach ($key in $($graphFolders.Keys)) {
        $graphPath = Join-Path -Path $reportPath -ChildPath $graphFolders[$key]
        $graphFolders[$key] = $graphPath
        New-Item -Path $graphPath -ItemType Directory
    }
    $graphFolders
}
function ConvertTo-Name {
    Param(
        [Parameter(Mandatory = $true)]
        [alias("ObjectList_DN", "ObjectList_DistinguishedName")]
        $objectListDN
    )
    $namesList = New-Object System.Collections.Generic.List[string]
    $objectListDN | ForEach-Object {

        if ($($_.contains("/"))) {
            $namesList.Add($($_.split("/"))[1])
        }
        else {
            $namesList.Add($($_ | Select-Object @{Name = 'Name'; expression = { $($_.split(',')[0]).split('=')[1] } }).Name)
        }
    }
    $namesList
}
##########################################################################################
#NOTES
####
#Local security groups - tworzone gdy tworzy się AD
#category = security/distribution
#scope=universal/global/domain_local
#builtin=tworzone przy starcie AD
##########################################################################################
function Invoke-ADAudit
{

$reportGraphFolders = Get-ReportFolders -BasePath $basePath -GraphFoldersHashtable $graphFolders

$reportFilePath = Join-Path -Path $basePath -ChildPath "report.docx"
$reportFile = New-WordDocument $reportFilePath


Add-WordText -WordDocument $reportFile -Text 'Raport z Active Directory' -FontSize 28 -FontFamily 'Calibri Light' -Supress $True
Add-WordPageBreak -WordDocument $reportFile -Supress $true

#region TOC #########################################################################################################

Add-WordTOC -WordDocument $reportFile -Title "Spis treści" -Supress $true

Add-WordPageBreak -WordDocument $reportFile -Supress $true

#endregion TOC ########################################################################################################

#region OU ############################################################################################################
Add-WordText -WordDocument $reportFile -HeadingType Heading1 -Text 'Spis jednostek organizacyjnych' -Supress $true
Add-WordText -WordDocument $reportFile -Text 'Ta część zawiera spis jednostek organizacyjnych wraz z informacjami o każdej z nich' -Supress $True

$ous = Get-OUInformation
foreach ($ou in $ous) 
{
    Add-WordText -WordDocument $reportFile -HeadingType Heading2 -Text $($ou.Name) -Supress $true
    
    Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "$($ou.Name) Information" -Supress $true
    Add-WordTable -WordDocument $reportFile -DataTable $ou -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle $($ou.Name) -Transpose  -Supress $True
    
    Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "$($ou.Name) Graph" -Supress $true 

    $ouTMP = $(Get-ADOrganizationalUnit -Filter "*" -SearchBase $($ou.DistinguishedName) -SearchScope OneLevel).Name
    if ($null -eq $ouTMP) {
        Add-WordText -WordDocument $reportFile -Text "No Leafs" -Supress $true     
    }
    else {
        $imagePath = Get-GraphImage -GraphRoot $($ou.Name) -GraphLeaf $ouTMP  -BasePathToGraphImage $($reportGraphFolders.OU)
        Add-WordPicture -WordDocument $reportFile -ImagePath $imagePath -Alignment center -ImageWidth 600 -Supress $True
    }

    #ACL
    $ouACL = Get-OUACL -OU $($ou.DistinguishedName)
    
    Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "$($ou.Name) Permissions" -Supress $true 
    Add-WordTable -WordDocument $reportFile -DataTable $($ouACL | Select-Object -Property * -ExcludeProperty ACLs) -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle "OU Options" -Transpose -Supress $true
    Add-WordText -WordDocument $reportFile -Text "" -Supress $true
    
    $($ouACL.ACLs) | ForEach-Object {
        Add-WordTable -WordDocument $reportFile -DataTable $($_) -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle "$($($_).IdentityReference) Permissions" -Transpose -Supress $true
        Add-WordText -WordDocument $reportFile -Text "" -Supress $true
    }
}

Add-WordText -WordDocument $reportFile -Text "OrganizationalUnit Tables"  -HeadingType Heading2 -Supress $true
Add-WordText -WordDocument $reportFile -Text "Tabela miast i Państw"  -HeadingType Heading3 -Supress $true
$ouTable = $($ous | Select-Object Name, StreetAddress, PostalCode, City, State, Country)
Add-WordTable -WordDocument $reportFile -DataTable $ouTable -Design ColorfulGridAccent1 -Supress $True #-Verbose

Add-WordText -WordDocument $reportFile -Text "OrganizationalUnit Lists"  -HeadingType Heading2 -Supress $true

Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "Ostatnie 10 zmienionych jednostek organizacyjnych" -Supress $true
$list = $($($ous | Select-Object whenChanged, Name | Sort-Object -Descending whenChanged | Select-Object -First 10) | Select-Object @{Name = "OUName"; Expression = { "$($_.Name) - $($_.whenChanged)" } }).OUName
Add-WordList -WordDocument $reportFile -ListType Numbered -ListData $list -Supress $true -Verbose

Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "Ostatnie 10 utworzonych jednostek organizacyjnych" -Supress $true
$list = $($($ous | Select-Object whenCreated, Name | Sort-Object -Descending whenCreated | Select-Object -First 10) | Select-Object @{Name = "OUName"; Expression = { "$($_.Name) - $($_.whenCreated)" } }).OUName
Add-WordList -WordDocument $reportFile -ListType Numbered -ListData $list -Supress $true -Verbose

#endregion OU #####################################################################################################

#region GROUPS#####################################################################################################
Add-WordText -WordDocument $reportFile -Text 'Spis Grup' -HeadingType Heading1 -Supress $true
Add-WordText -WordDocument $reportFile -Text 'Jest to dokumentacja domeny ActiveDirectory przeprowadzona w domena.local. Wszytskie informacje są tajne' -Supress $True 

$groups = Get-GROUPInformation

Add-WordText -WordDocument $reportFile -Text "DomainLocal groups"  -HeadingType Heading2 -Supress $true

$groupObjects=$groups | Where-Object {$_.GroupType -band 1 }

foreach ($group in $groupObjects) 
{
    Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text $($group.Name) -Supress $true
    
    Add-WordText -WordDocument $reportFile -HeadingType Heading4 -Text "$($group.Name) Information" -Supress $true
    Add-WordTable -WordDocument $reportFile -DataTable $group -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle $($group.Name) -Transpose -Supress $True
    
    Add-WordText -WordDocument $reportFile -HeadingType Heading4 -Text "$($group.Name) Graph" -Supress $true

    if ($null -eq $($group.Members)) 
    {
        Add-WordText -WordDocument $reportFile -Text "No Leafs" -Supress $true    
    }
    else 
    {
        $groupMembersTMP = ConvertTo-Name -ObjectList_DN $($group.Members)
        $imagePath = Get-GraphImage -GraphRoot $($group.Name) -GraphLeaf $groupMembersTMP -pathToImage $($reportGraphFolders.GROUP)
        Add-WordPicture -WordDocument $reportFile -ImagePath $imagePath -Alignment center -ImageWidth 600 -Supress $True
    }
}


Add-WordText -WordDocument $reportFile -Text "Security Groups"  -HeadingType Heading2 -Supress $true

$groupObjects=$groups | Where-Object { (-not($_.GroupType -band 1)) -and ($_.GroupCategory -eq "Security") }
 
foreach ($group in $groupObjects) 
{
    Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text $($group.Name) -Supress $true
    
    Add-WordText -WordDocument $reportFile -HeadingType Heading4 -Text "$($group.Name) Information" -Supress $true
    Add-WordTable -WordDocument $reportFile -DataTable $group -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle $($group.Name) -Transpose -Supress $True
    
    Add-WordText -WordDocument $reportFile -HeadingType Heading4 -Text "$($group.Name) Graph" -Supress $true

    if ($null -eq $($group.Members)) 
    {
        Add-WordText -WordDocument $reportFile -Text "No Leafs" -Supress $true    
    }
    else 
    {
        $groupMembersTMP = ConvertTo-Name -ObjectList_DN $($group.Members)
        $imagePath = Get-GraphImage -GraphRoot $($group.Name) -GraphLeaf $groupMembersTMP -pathToImage $($reportGraphFolders.GROUP)
        Add-WordPicture -WordDocument $reportFile -ImagePath $imagePath -Alignment center -ImageWidth 600 -Supress $True
    }
}


Add-WordText -WordDocument $reportFile -Text "Distribution Groups"  -HeadingType Heading2 -Supress $true

$groupObjects=$groups | Where-Object {$_.GroupCategory -eq "Distribution" }
 
foreach ($group in $groupObjects) 
{
    Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text $($group.Name) -Supress $true
    
    Add-WordText -WordDocument $reportFile -HeadingType Heading4 -Text "$($group.Name) Information" -Supress $true
    Add-WordTable -WordDocument $reportFile -DataTable $group -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle $($group.Name) -Transpose -Supress $True
    
    Add-WordText -WordDocument $reportFile -HeadingType Heading4 -Text "$($group.Name) Graph" -Supress $true

    if ($null -eq $($group.Members)) 
    {
        Add-WordText -WordDocument $reportFile -Text "No Leafs" -Supress $true    
    }
    else 
    {
        $groupMembersTMP = ConvertTo-Name -ObjectList_DN $($group.Members)
        $imagePath = Get-GraphImage -GraphRoot $($group.Name) -GraphLeaf $groupMembersTMP -pathToImage $($reportGraphFolders.GROUP)
        Add-WordPicture -WordDocument $reportFile -ImagePath $imagePath -Alignment center -ImageWidth 600 -Supress $True
    }
}

Add-WordText -WordDocument $reportFile -Text "Group Charts"  -HeadingType Heading2 -Supress $true

#TEST
$chart = $groups | Group-Object GroupCategory | Select-Object Name, Count
Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "Wykresy grup dystrybucyjnych/zabezpieczeń" -Supress $true
Add-WordBarChart -WordDocument $reportFile -ChartName 'Stosunek liczby grup zabezpieczeń do grup dystrybucyjnych'-ChartLegendPosition Bottom -ChartLegendOverlay $false -Names $([array]$chart.Name) -Values $([array]$chart.Count) -BarDirection Column
#TEST
$chart = $groups | Group-Object GroupScope | Select-Object Name, Count
Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "Wykresy grup lokalnych/globalnych/uniwersalnych" -Supress $true
Add-WordBarChart -WordDocument $reportFile -ChartName 'Stosunek liczby grup lokalnych, globalnych,uniwersalnych'-ChartLegendPosition Bottom -ChartLegendOverlay $false -Names $([array]$chart.Name) -Values $([array]$chart.Count) -BarDirection Column

Add-WordText -WordDocument $reportFile -Text "Group Lists"  -HeadingType Heading2 -Supress $true

$list = $($($groups | Select-Object whenChanged, Name | Sort-Object -Descending whenChanged | Select-Object -First 10) | Select-Object @{Name = "GroupName"; Expression = { "$($_.Name) - $($_.whenChanged)" } }).GroupName
Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "Ostatnie 10 zmienionych grup" -Supress $true
Add-WordList -WordDocument $reportFile -ListType Numbered -ListData $list -Supress $true -Verbose

$list = $($($groups | Select-Object whenCreated, Name | Sort-Object -Descending whenCreated | Select-Object -First 10) | Select-Object @{Name = "GroupName"; Expression = { "$($_.Name) - $($_.whenCreated)" } }).GroupName
Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "Ostatnie 10 utworzonych grup" -Supress $true
Add-WordList -WordDocument $reportFile -ListType Numbered -ListData $list -Supress $true -Verbose

$list = $($groups | Where-Object { $_.Members.Count -eq 0 }).Name
Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "Grupy puste" -Supress $true
Add-WordList -WordDocument $reportFile -ListType Numbered -ListData $list -Supress $true -Verbose

Add-WordText -WordDocument $reportFile -Text "Group Tables"  -HeadingType Heading2 -Supress $true
Add-WordText -WordDocument $reportFile -Text "Tabela grup"  -HeadingType Heading3 -Supress $true

$groupTable = $groups | Group-Object GroupScope | ForEach-Object {
    $categories = $_.Group | Group-Object GroupCategory -AsHashtable -AsString

    [PSCustomObject]@{
        GroupName    = $_.Name
        Security     = $categories['Security'].Count
        Distribution = $categories['Distribution'].Count
    }
}

Add-WordTable -WordDocument $reportFile -DataTable $groupTable -Design ColorfulGridAccent1 -Supress $True #-Verbose

#TODO:Group Graphs
#endregion GROUPS#####################################################################################################

#region USERS#####################################################################################################
Add-WordText -WordDocument $reportFile -Text 'Spis Użytkowników' -HeadingType Heading1 -Supress $true
Add-WordText -WordDocument $reportFile -Text 'Ta część zawiera spis użytkowników domeny' -Supress $True 

$users = Get-USERInformation

foreach ($user in $users) 
{
    Add-WordText -WordDocument $reportFile -HeadingType Heading2 -Text $($user.Name) -Supress $true
    
    Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "$($user.Name) Information" -Supress $true
    Add-WordTable -WordDocument $reportFile -DataTable $user -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle $($user.Name) -Transpose -Supress $true
 
    #MemberOf
    Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "$($user.Name) MemberOfGroup Graph" -Supress $true 

    if ($null -eq $($usert.MemberOf)) {
        Add-WordText -WordDocument $reportFile -Text "No Leafs" -Supress $true     
    }
    else {
        $memberOfTMP = ConvertTo-Name -ObjectList_DN $($user.MemberOf)
        $imagePath = Get-GraphImage -GraphRoot $($user.Name) -GraphLeaf $memberOfTMP  -BasePathToGraphImage $($reportGraphFolders.USERS)
        Add-WordPicture -WordDocument $reportFile -ImagePath $imagePath -Alignment center -ImageWidth 600 -Supress $True
    }

    #Manager
    Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "$($user.Name) DirectManager" -Supress $true 

    if ($null -eq $($user.Manager)) {
        Add-WordText -WordDocument $reportFile -Text "No Leafs" -Supress $true     
    }
    else {
        $managerTMP = ConvertTo-Name -ObjectList_DN $($user.Manager)
        $imagePath = Get-GraphImage -GraphRoot $managerTMP -GraphLeaf $($user.Name)  -BasePathToGraphImage $($reportGraphFolders.USERS)
        Add-WordPicture -WordDocument $reportFile -ImagePath $imagePath -Alignment center -ImageWidth 600 -Supress $True
    }

    #TODO:Create graph with full organisation manager and direct report
}


Add-WordText -WordDocument $reportFile -Text "Users Table"  -HeadingType Heading2 -Supress $true

Add-WordText -WordDocument $reportFile -Text "Tabela lokalizacji użytkowników"  -HeadingType Heading3 -Supress $true
$table = $($users | Select-Object Name, Department, City, Country)
Add-WordTable -WordDocument $reportFile -DataTable $table -Design ColorfulGridAccent5 -AutoFit Window -Supress $true

Add-WordText -WordDocument $reportFile -Text "Tabela bezpieczeństwa"  -HeadingType Heading3 -Supress $true
$table = $($users | Select-Object Name, CannotChangePassword, PasswordExpired, PasswordNeverExpires, PasswordNotRequired)
Add-WordTable -WordDocument $reportFile -DataTable $table -Design ColorfulGridAccent5 -AutoFit Window -Supress $true

Add-WordText -WordDocument $reportFile -Text "Users Lists"  -HeadingType Heading2 -Supress $true

Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "Ostatnie 10 zmienionych użytkowników" -Supress $true
$list = $($($users | Select-Object whenChanged, Name | Sort-Object -Descending whenChanged | Select-Object -First 10) | Select-Object @{Name = "UserName"; Expression = { "$($_.Name) - $($_.whenChanged)" } }).UserName
Add-WordList -WordDocument $reportFile -ListType Numbered -ListData $list -Supress $true -Verbose

Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "Ostatnie 10 utworzonych użytkowników" -Supress $true
$list = $($($users | Select-Object whenCreated, Name | Sort-Object -Descending whenCreated | Select-Object -First 10) | Select-Object @{Name = "UserName"; Expression = { "$($_.Name) - $($_.whenCreated)" } }).UserName
Add-WordList -WordDocument $reportFile -ListType Numbered -ListData $list -Supress $true -Verbose

Add-WordText -WordDocument $reportFile -Text "User Charts"  -HeadingType Heading2 -Supress $true

$chart = $users | Group-Object Enabled | Select-Object Name, Count
Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "Wykresy kont wyłączonych/włączonych" -Supress $true
Add-WordBarChart -WordDocument $reportFile -ChartName 'Stosunek liczby kont wyłączonych i włączonych'-ChartLegendPosition Bottom -ChartLegendOverlay $false -Names $([array]$chart.Name) -Values $([array]$chart.Count) -BarDirection Column

#TEST
$chart = $users | Group-Object Office | Select-Object Name , @{Name = "Values"; Expression = {[math]::round(((($_.Count) / $users.Count) * 100), 2)} } | Where-Object { $_.Values -ge 1 } | Sort-Object -Descending Values
Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "Wykres biur w przekroju firmy" -Supress $true 
Add-WordPieChart -WordDocument $reportFile -ChartName 'Stosunek liczby stanowisk'-ChartLegendPosition Bottom -ChartLegendOverlay $false -Names $([array]$chart.Name) -Values $([array]$chart.Values)
#TEST
$chart = $users | Group-Object Title | Select-Object Name, @{Name = "Values"; Expression = { [math]::round(((($_.Count) / $users.Count) * 100), 2)} } | Where-Object { $_.Values -ge 1 } | Sort-Object -Descending Values
Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "Wykres stanowisk w przekroju firmy" -Supress $true
Add-WordPieChart -WordDocument $reportFile -ChartName 'Stosunek liczby stanowisk'-ChartLegendPosition Bottom -ChartLegendOverlay $false -Names $([array]$chart.Name) -Values $([array]$chart.Values)
#TEST
$chart = $users | Group-Object Department | Select-Object Name, @{Name = "Values"; Expression = {[math]::round(((($_.Count) / $users.Count) * 100), 2)} } | Where-Object { $_.Values -ge 1 } | Sort-Object -Descending Values
Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "Wykres departamentów w przekroju firmy" -Supress $true
Add-WordPieChart -WordDocument $reportFile -ChartName 'Stosunek liczby stanowisk'-ChartLegendPosition Bottom -ChartLegendOverlay $false -Names $([array]$chart.Name) -Values $([array]$chart.Values)

#endregion USERS#####################################################################################################

#region GPO############################################################################################################
Add-WordText -WordDocument $reportFile -HeadingType Heading1 -Text 'Spis Polis Grup' -Supress $true
Add-WordText -WordDocument $reportFile -Text 'Tutaj znajduje się opis polis grup. Blok nie pokazuje polis podłączonych do SITE' -Supress $True

$groupPolicyObjects = Get-GPOPolicy

foreach ($gpoPolicyObject in $groupPolicyObjects) 
{
    Add-WordText -WordDocument $reportFile -HeadingType Heading2 -Text $($gpoPolicyObject.Name) -Supress $true
    
    Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "$($gpoPolicyObject.Name) Information" -Supress $true
    Add-WordTable -WordDocument $reportFile -DataTable $gpoPolicyObject -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle $($gpoPolicyObject.Name) -Transpose -Supress $true
 
    Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "$($gpoPolicyObject.Name) Graph" -Supress $true
    
    if ($null -eq $($gpoPolicyObject.Links)) {
        Add-WordText -WordDocument $reportFile -Text "No Leafs" -Supress $true    
    }
    else {
        $linksTMP = ConvertTo-Name -ObjectList_DN $($gpoPolicyObject.Links)
        $imagePath = Get-GraphImage -GraphRoot $($gpoPolicyObject.Name) -GraphLeaf $linksTMP -pathToImage $($reportGraphFolders.GPO)
        Add-WordPicture -WordDocument $reportFile -ImagePath $imagePath -Alignment center -ImageWidth 600 -Supress $True
    }

    Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "$($gpoPolicyObject.Name) Permissions" -Supress $true
    
    #ACL
    $gpoACL = $(Get-GPOAcl -GroupPolicyObject $gpoPolicyObject).ACLs
    $gpoACL | ForEach-Object {
        Add-WordTable -WordDocument $reportFile -DataTable $($_) -Design ColorfulGridAccent5 -AutoFit Window -Supress $true -Transpose
        Add-WordText -WordDocument $reportFile -Text "" -Supress $true
    }
}


Add-WordText -WordDocument $reportFile -Text "Group Policy Lists"  -HeadingType Heading2 -Supress $true 

Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "Ostatnie 10 utworzonych jednostek organizacyjnych" -Supress $true
$list = $($($groupPolicyObjects | Select-Object ModificationTime, Name | Sort-Object -Descending ModificationTime | Select-Object -First 10) | Select-Object @{Name = "GPOName"; Expression = { "$($_.Name) - $($_.ModificationTime)" } }).GPOName
Add-WordList -WordDocument $reportFile -ListType Numbered -ListData $list -Supress $true -Verbose

Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "Ostatnie 10 utworzonych polis grup" -Supress $true
$list = $($($groupPolicyObjects | Select-Object CreationTime, Name | Sort-Object -Descending CreationTime | Select-Object -First 10) | Select-Object @{Name = "GPOName"; Expression = { "$($_.Name) - $($_.CreationTime)" } }).GPOName
Add-WordList -WordDocument $reportFile -ListType Numbered -ListData $list -Supress $true -Verbose

Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "Polisy grup nieprzypisane" -Supress $true
$list = $($groupPolicyObjects | Where-Object { $_.Links.Count -eq 0 }).Name
Add-WordList -WordDocument $reportFile -ListType Numbered -ListData $list -Supress $true -Verbose

Add-WordText -WordDocument $reportFile -Text "GroupPolicy Tables"  -HeadingType Heading2 -Supress $true

Add-WordText -WordDocument $reportFile -Text "Tabela polis grup"  -HeadingType Heading3 -Supress $true
$gpoTable = $($groupPolicyObjects | Select-Object Name, HasComputerSettings, HasUserSettings, UserEnabled, ComputerEnabled, ComputerSettings, UserSettings)
Add-WordTable -WordDocument $reportFile -DataTable $gpoTable -Design ColorfulGridAccent1 -Supress $True #-Verbose

#endregion GPO################################################################################################

#region FGPP##################################################################################################

Add-WordText -WordDocument $reportFile -HeadingType Heading1 -Text 'Spis Fine Grained Password Policies' -Supress $true
Add-WordText -WordDocument $reportFile -Text 'Tutaj znajduje się opis obiektów Fine Grained Password Policies' -Supress $True

$fgpps = Get-FineGrainedPolicies
foreach ($fgpp in $fgpps) {
    Add-WordText -WordDocument $reportFile -HeadingType Heading2 -Text $($fgpp.Name) -Supress $true
    
    Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "$($fgpp.Name) Information" -Supress $true
    Add-WordTable -WordDocument $reportFile -DataTable $fgpp -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle $($fgpp.Name) -Transpose -Supress $true

    Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "$($fgpp.Name) is applied to" -Supress $true
    
    if ($null -eq $($fgpp.'Applies To')) {
        Add-WordText -WordDocument $reportFile -Text "No Leafs" -Supress $true    
    }
    else {
        $fgppAplliedTMP = ConvertTo-Name -ObjectList_DN $($fgpp.'Applies To')
        $imagePath = Get-GraphImage -GraphRoot $($fgpp.Name) -GraphLeaf $fgppAplliedTMP -pathToImage $reportGraphFolders.FGPP
        Add-WordPicture -WordDocument $reportFile -ImagePath $imagePath -Alignment center -ImageWidth 600 -Supress $True
    }
    

}
#endregion FGPP###############################################################################################
#TODO:ComputerReport
#TODO: Flatten ACL

#Get-ComputerInformation
##############################################################################################################
Save-WordDocument $reportFile -Supress $true -Language "pl-PL" -Verbose #-OpenDocument
Invoke-Item -Path $reportFilePath
}
Invoke-ADAudit