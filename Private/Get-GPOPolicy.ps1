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