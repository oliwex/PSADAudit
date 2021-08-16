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