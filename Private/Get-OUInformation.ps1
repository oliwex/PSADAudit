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