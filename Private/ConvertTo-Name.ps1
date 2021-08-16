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