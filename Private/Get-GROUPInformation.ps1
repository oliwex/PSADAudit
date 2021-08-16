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