#TODO: Big graph representing full company/representing OU structure?
#TODO:Create ACLS based on adsi edit for every element
##########################################################################################
#                                GLOBAL VARIABLES                                        #
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
$basePath=New-Workplace
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

    $ouLeaf = $(Get-ADOrganizationalUnit -Filter "*" -SearchBase $($ou.DistinguishedName) -SearchScope OneLevel).Name
    $ouRoot=$($($ou.DistinguishedName) -split ',*..=')[2]
    if (($null -eq $ouLeaf) -and ($null -eq $ouRoot))
    {
        Add-WordText -WordDocument $reportFile -Text "$($ou.Name) do not have above and below elements" -Supress $true
    }
    else
    { 
        $imagePath = Get-GraphImage -GraphRoot $ouRoot -GraphMiddle $($ou.Name) -GraphLeaf $ouLeaf  -BasePathToGraphImage $($reportGraphFolders.OU)
        Add-WordPicture -WordDocument $reportFile -ImagePath $imagePath -Alignment center -ImageWidth 600 -Supress $True
    }
   
    #ACL
    $ouACL = Get-OUAcl -OU $($ou.DistinguishedName)
    
    Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "$($ou.Name) Permissions" -Supress $true 
    Add-WordTable -WordDocument $reportFile -DataTable $($ouACL | Select-Object -Property * -ExcludeProperty ACLs) -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle "OU Options" -Transpose -Supress $true
    Add-WordText -WordDocument $reportFile -Text "" -Supress $true
    
    Add-WordTable -WordDocument $reportFile -DataTable $($ouACL.ACLs) -Design MediumShading1Accent5 -AutoFit Window  -Supress $true
    Add-WordText -WordDocument $reportFile -Text "" -Supress $true
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


    $groupLeafTMP = $group.Members | ForEach-Object { $(($_ -split ',*..=')[1]) }
    $groupRootTMP = $group.MemberOf | ForEach-Object { $(($_ -split ',*..=')[1]) }
    if (($null -eq $groupLeafTMP) -and ($null -eq $groupRootTMP))
    {
        Add-WordText -WordDocument $reportFile -Text "$($group.Name) do not have above and below elements" -Supress $true
    }
    else
    {
        $imagePath = Get-GraphImage -GraphRoot $groupRootTMP -GraphMiddle $($group.Name) -GraphLeaf $groupLeafTMP -pathToImage $($reportGraphFolders.GROUP)
        Add-WordPicture -WordDocument $reportFile -ImagePath $imagePath -Alignment center -ImageWidth 600 -Supress $True
    }

    #ACL
    $groupACL=Get-GROUPAcl -GROUP_ACL $($group.DistinguishedName)

    Add-WordText -WordDocument $reportFile -HeadingType Heading4 -Text "$($group.Name) Permissions" -Supress $true 
    Add-WordTable -WordDocument $reportFile -DataTable $($groupACL | Select-Object -Property * -ExcludeProperty ACLs) -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle "OU Options" -Transpose -Supress $true
    Add-WordText -WordDocument $reportFile -Text "" -Supress $true
    
    Add-WordTable -WordDocument $reportFile -DataTable $($groupACL.ACLs) -Design MediumShading1Accent5 -AutoFit Window  -Supress $true
    Add-WordText -WordDocument $reportFile -Text "" -Supress $true
}


Add-WordText -WordDocument $reportFile -Text "Security Groups"  -HeadingType Heading2 -Supress $true

$groupObjects=$groups | Where-Object { (-not($_.GroupType -band 1)) -and ($_.GroupCategory -eq "Security") }
 
foreach ($group in $groupObjects) 
{
    Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text $($group.Name) -Supress $true
    
    Add-WordText -WordDocument $reportFile -HeadingType Heading4 -Text "$($group.Name) Information" -Supress $true
    Add-WordTable -WordDocument $reportFile -DataTable $group -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle $($group.Name) -Transpose -Supress $True
    
    Add-WordText -WordDocument $reportFile -HeadingType Heading4 -Text "$($group.Name) Graph" -Supress $true

    $groupLeafTMP = $group.Members | ForEach-Object { $(($_ -split ',*..=')[1]) }
    $groupRootTMP = $group.MemberOf | ForEach-Object { $(($_ -split ',*..=')[1]) }
    if (($null -eq $groupLeafTMP) -and ($null -eq $groupRootTMP))
    {
        Add-WordText -WordDocument $reportFile -Text "$($group.Name) do not have above and below elements" -Supress $true
    }
    else
    {
        $imagePath = Get-GraphImage -GraphRoot $groupRootTMP -GraphMiddle $($group.Name) -GraphLeaf $groupLeafTMP -pathToImage $($reportGraphFolders.GROUP)
        Add-WordPicture -WordDocument $reportFile -ImagePath $imagePath -Alignment center -ImageWidth 600 -Supress $True
    }

    #ACL
    $groupACL=Get-GROUPAcl -GROUP_ACL $($group.DistinguishedName)

    Add-WordText -WordDocument $reportFile -HeadingType Heading4 -Text "$($group.Name) Permissions" -Supress $true 
    Add-WordTable -WordDocument $reportFile -DataTable $($groupACL | Select-Object -Property * -ExcludeProperty ACLs) -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle "OU Options" -Transpose -Supress $true
    Add-WordText -WordDocument $reportFile -Text "" -Supress $true
    
    Add-WordTable -WordDocument $reportFile -DataTable $($groupACL.ACLs) -Design MediumShading1Accent5 -AutoFit Window  -Supress $true
    Add-WordText -WordDocument $reportFile -Text "" -Supress $true
}

Add-WordText -WordDocument $reportFile -Text "Distribution Groups"  -HeadingType Heading2 -Supress $true

$groupObjects=$groups | Where-Object {$_.GroupCategory -eq "Distribution" }
 
foreach ($group in $groupObjects) 
{
    Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text $($group.Name) -Supress $true
    
    Add-WordText -WordDocument $reportFile -HeadingType Heading4 -Text "$($group.Name) Information" -Supress $true
    Add-WordTable -WordDocument $reportFile -DataTable $group -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle $($group.Name) -Transpose -Supress $True
    
    Add-WordText -WordDocument $reportFile -HeadingType Heading4 -Text "$($group.Name) Graph" -Supress $true

    $groupLeafTMP = $group.Members | ForEach-Object { $(($_ -split ',*..=')[1]) }
    $groupRootTMP = $group.MemberOf | ForEach-Object { $(($_ -split ',*..=')[1]) }
    if (($null -eq $groupLeafTMP) -and ($null -eq $groupRootTMP))
    {
        Add-WordText -WordDocument $reportFile -Text "$($group.Name) do not have above and below elements" -Supress $true
    }
    else
    {
        $imagePath = Get-GraphImage -GraphRoot $groupRootTMP -GraphMiddle $($group.Name) -GraphLeaf $groupLeafTMP -pathToImage $($reportGraphFolders.GROUP)
        Add-WordPicture -WordDocument $reportFile -ImagePath $imagePath -Alignment center -ImageWidth 600 -Supress $True
    }

    #ACL
    $groupACL=Get-GROUPAcl -GROUP_ACL $($group.DistinguishedName)

    Add-WordText -WordDocument $reportFile -HeadingType Heading4 -Text "$($group.Name) Permissions" -Supress $true 
    Add-WordTable -WordDocument $reportFile -DataTable $($groupACL | Select-Object -Property * -ExcludeProperty ACLs) -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle "OU Options" -Transpose -Supress $true
    Add-WordText -WordDocument $reportFile -Text "" -Supress $true
    
    Add-WordTable -WordDocument $reportFile -DataTable $($groupACL.ACLs) -Design MediumShading1Accent5 -AutoFit Window  -Supress $true
    Add-WordText -WordDocument $reportFile -Text "" -Supress $true
}

Add-WordText -WordDocument $reportFile -Text "Group Charts"  -HeadingType Heading2 -Supress $true

$chart = $groups | Group-Object GroupCategory | Select-Object Name, @{Name="Values";Expression={$_.Count}}
Add-WordChart -CType "Barchart" -CData $chart -STitle "Wykresy grup dystrybucyjnych/zabezpieczeń" -CTitle "Stosunek liczby grup zabezpieczeń do grup dystrybucyjnych"

$chart = $groups | Group-Object GroupScope | Select-Object Name, @{Name="Values";Expression={$_.Count}}
Add-WordChart -CType "Barchart" -CData $chart -STitle "Wykresy grup lokalnych/globalnych/uniwersalnych" -CTitle "Stosunek liczby grup lokalnych, globalnych,uniwersalnych"

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

    if ($null -eq $($user.MemberOf)) {
        Add-WordText -WordDocument $reportFile -Text "$($user.Name) do not have below elements" -Supress $true   
    }
    else {
        $memberOfTMP = $user.MemberOf | ForEach-Object { $(($_ -split ',*..=')[1]) }
        $imagePath = Get-GraphImage -GraphRoot $null -GraphMiddle $($user.Name) -GraphLeaf $memberOfTMP  -BasePathToGraphImage $($reportGraphFolders.USERS)
        Add-WordPicture -WordDocument $reportFile -ImagePath $imagePath -Alignment center -ImageWidth 600 -Supress $True
    }

    #Manager
    Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "$($user.Name) DirectManager" -Supress $true 
    $managerTMP = $user.Manager| ForEach-Object { $(($_ -split ',*..=')[1]) } 
    $directReportsTMP = $user.DirectReports | ForEach-Object { $(($_ -split ',*..=')[1]) }
    if (($null -eq $managerTMP) -and ($null -eq $directReportsTMP))
    {
        Add-WordText -WordDocument $reportFile -Text "$($user.Name) do not have above and below elements" -Supress $true    
    }
    else
    {
        $imagePath = Get-GraphImage -GraphRoot $managerTMP -GraphMiddle $($user.Name) -GraphLeaf $directReportsTMP  -BasePathToGraphImage $($reportGraphFolders.USERS)
        Add-WordPicture -WordDocument $reportFile -ImagePath $imagePath -Alignment center -ImageWidth 600 -Supress $True
    }
    #TODO:Create graph with full organisation manager and direct report

    #ACL
    $userACL = Get-USERAcl -USER_ACL $($user.DistinguishedName)

    Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "$($user.Name) Permissions" -Supress $true 
    Add-WordTable -WordDocument $reportFile -DataTable $($userACL | Select-Object -Property * -ExcludeProperty ACLs) -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle "User Options" -Transpose -Supress $true
    Add-WordText -WordDocument $reportFile -Text "" -Supress $true
    
    Add-WordTable -WordDocument $reportFile -DataTable $($userACL.ACLs) -Design MediumShading1Accent5 -AutoFit Window  -Supress $true
    Add-WordText -WordDocument $reportFile -Text "" -Supress $true

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

$chart = $users | Group-Object Enabled | Select-Object Name, @{Name="Values";Expression={$_.Count}}
Add-WordChart -CType "Barchart" -CData $chart -STitle "Wykresy kont wyłączonych/włączonych" -CTitle "Stosunek liczby kont wyłączonych i włączonych"

$chart = $users | Group-Object Office | Select-Object Name , @{Name = "Values"; Expression = {[math]::round(((($_.Count) / $users.Count) * 100), 2)} } | Where-Object { $_.Values -ge 1 } | Sort-Object -Descending Values
Add-WordChart -CType "Piechart" -CData $chart -STitle "Wykres biur w przekroju firmy" -CTitle "Stosunek liczby stanowisk"

$chart = $users | Group-Object Title | Select-Object Name, @{Name = "Values"; Expression = { [math]::round(((($_.Count) / $users.Count) * 100), 2)} } | Where-Object { $_.Values -ge 1 } | Sort-Object -Descending Values
Add-WordChart -CType "Piechart" -CData $chart -STitle "Wykres stanowisk w przekroju firmy" -CTitle "Stosunek liczby stanowisk"

$chart = $users | Group-Object Department | Select-Object Name, @{Name = "Values"; Expression = {[math]::round(((($_.Count) / $users.Count) * 100), 2)} } | Where-Object { $_.Values -ge 1 } | Sort-Object -Descending Values
Add-WordChart -CType "Piechart" -CData $chart -STitle "Wykres departamentów w przekroju firmy" -CTitle "Stosunek liczby stanowisk"

#endregion USERS#####################################################################################################

#region GPO############################################################################################################
Add-WordText -WordDocument $reportFile -HeadingType Heading1 -Text 'Spis Polis Grup' -Supress $true
Add-WordText -WordDocument $reportFile -Text 'Tutaj znajduje się opis polis grup. Blok nie pokazuje polis podłączonych do SITE' -Supress $True

$groupPolicyObjects = Get-GPO -Domain $($Env:USERDNSDOMAIN) -All 

$groupPolicyObjectsList = foreach ($groupPolicyObject in $groupPolicyObjects) 
{
    $gpoObject = Get-GPOPolicy -GroupPolicy $groupPolicyObject
    
    Add-WordText -WordDocument $reportFile -HeadingType Heading2 -Text $($gpoObject.Name) -Supress $true
    
    Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "$($gpoObject.Name) Information" -Supress $true
    Add-WordTable -WordDocument $reportFile -DataTable $gpoObject -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle $($gpoObject.Name) -Transpose -Supress $true
    Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "$($gpoObject.Name) Graph" -Supress $true
    
    
    if ($null -eq $($gpoObject.Links)) 
    {
        Add-WordText -WordDocument $reportFile -Text "$($gpoObject.Name) do not have below elements" -Supress $true      
    }
    else 
    {
        $linksTMP = $gpoObject.Links.split(";")
        $imagePath = Get-GraphImage -GraphRoot $null -GraphMiddle $($gpoObject.Name) -GraphLeaf $linksTMP -pathToImage $($reportGraphFolders.GPO)
        Add-WordPicture -WordDocument $reportFile -ImagePath $imagePath -Alignment center -ImageWidth 600 -Supress $True
    }
    
    #ACL
    Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "$($gpoObject.Name) Permissions Simple" -Supress $true
    $gpoObjectACL = Get-GPOAclSimple -GroupPolicy $groupPolicyObject
    
    $gpoObjectACL.ACL | ForEach-Object {
        Add-WordTable -WordDocument $reportFile -DataTable $($_) -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle "Permissions" -Transpose -Supress $true
        Add-WordText -WordDocument $reportFile -Text "" -Supress $true 
    }

    #ACL
    $pathACL = Get-GPOAclExtended -GPO_ACL $($groupPolicyObject.Path)

    Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "$($gpoObject.Name) Permissions Extended" -Supress $true 
    Add-WordTable -WordDocument $reportFile -DataTable $($pathACL | Select-Object -Property * -ExcludeProperty ACLs) -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle "GPO Options" -Transpose -Supress $true
    Add-WordText -WordDocument $reportFile -Text "" -Supress $true
    
    Add-WordTable -WordDocument $reportFile -DataTable $($pathACL.ACLs) -Design MediumShading1Accent5 -AutoFit Window  -Supress $true
    Add-WordText -WordDocument $reportFile -Text "" -Supress $true

    $gpoObject
}


Add-WordText -WordDocument $reportFile -Text "Group Policy Lists"  -HeadingType Heading2 -Supress $true 

Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "Ostatnie 10 utworzonych jednostek organizacyjnych" -Supress $true
$list = $($($groupPolicyObjectsList | Select-Object ModificationTime, Name | Sort-Object -Descending ModificationTime | Select-Object -First 10) | Select-Object @{Name = "GPOName"; Expression = { "$($_.Name) - $($_.ModificationTime)" } }).GPOName
Add-WordList -WordDocument $reportFile -ListType Numbered -ListData $list -Supress $true -Verbose

Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "Ostatnie 10 utworzonych polis grup" -Supress $true
$list = $($($groupPolicyObjectsList | Select-Object CreationTime, Name | Sort-Object -Descending CreationTime | Select-Object -First 10) | Select-Object @{Name = "GPOName"; Expression = { "$($_.Name) - $($_.CreationTime)" } }).GPOName
Add-WordList -WordDocument $reportFile -ListType Numbered -ListData $list -Supress $true -Verbose

Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "Polisy grup nieprzypisane" -Supress $true
$list = $($groupPolicyObjectsList | Where-Object { $_.Links.Count -eq 0 }).Name
Add-WordList -WordDocument $reportFile -ListType Numbered -ListData $list -Supress $true -Verbose

Add-WordText -WordDocument $reportFile -Text "GroupPolicy Tables"  -HeadingType Heading2 -Supress $true

Add-WordText -WordDocument $reportFile -Text "Tabela polis grup 1"  -HeadingType Heading3 -Supress $true
$gpoTable = $($groupPolicyObjectsList | Select-Object Name, HasComputerSettings, HasUserSettings, ComputerSettings, UserSettings)
Add-WordTable -WordDocument $reportFile -DataTable $gpoTable -Design ColorfulGridAccent1 -Supress $True #-Verbose

Add-WordText -WordDocument $reportFile -Text "Tabela polis grup 2"  -HeadingType Heading3 -Supress $true
$gpoTable = $($groupPolicyObjectsList | Select-Object Name,UserEnabled, ComputerEnabled)
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
    
    if ($null -eq $($fgpp.'Applies To')) 
    {
        Add-WordText -WordDocument $reportFile -Text "$($fgpp.Name) do not have below elements" -Supress $true  
    }
    else 
    {
        $fgppAplliedTMP = $($fgpp.'Applies To').split(";")
        $imagePath = Get-GraphImage -GraphRoot $null -GraphMiddle $($fgpp.Name) -GraphLeaf $fgppAplliedTMP -pathToImage $reportGraphFolders.FGPP
        Add-WordPicture -WordDocument $reportFile -ImagePath $imagePath -Alignment center -ImageWidth 600 -Supress $True
    }
    

}
#endregion FGPP###############################################################################################

#region COMPUTERS#############################################################################################

Add-WordText -WordDocument $reportFile -HeadingType Heading1 -Text 'Spis Komputerów' -Supress $true
Add-WordText -WordDocument $reportFile -Text 'Tutaj znajduje się spis komputerów' -Supress $True

$computers=Get-ComputerInformation
foreach ($computer in $computers)
{
        Add-WordText -WordDocument $reportFile -HeadingType Heading2 -Text $($computer.Name) -Supress $true
    
        Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "$($computer.Name) Information" -Supress $true
        Add-WordTable -WordDocument $reportFile -DataTable $computer -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle $($computer.Name) -Transpose -Supress $true
        
        #MemberOf
        Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "$($computer.Name) MemberOfGroup Graph" -Supress $true 
        
        $computerLeafTMP = $computer.MemberOf| ForEach-Object { $(($_ -split ',*..=')[1]) }
        if ($null -eq $computerLeafTMP) 
        {
            Add-WordText -WordDocument $reportFile -Text "$($computer.Name) do not have below elements" -Supress $true     
        }
        else 
        {
            $imagePath = Get-GraphImage -GraphRoot $null -GraphMiddle $($computer.Name) -GraphLeaf $computerLeafTMP  -BasePathToGraphImage $($reportGraphFolders.COMPUTERS)
            Add-WordPicture -WordDocument $reportFile -ImagePath $imagePath -Alignment center -ImageWidth 600 -Supress $True
        }

        #ManagedBy
        Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "$($computer.Name) ManagedBy" -Supress $true 
        
        $managerTMP = $computer.ManagedBy | ForEach-Object { $(($_ -split ',*..=')[1]) }
        if ($null -eq $managerTMP) 
        {
            Add-WordText -WordDocument $reportFile -Text "$($computer.Name) do not have above elements" -Supress $true       
        }
        else 
        {
            $imagePath = Get-GraphImage -GraphRoot $null -GraphMiddle $managerTMP -GraphLeaf $($computer.Name)  -BasePathToGraphImage $($reportGraphFolders.COMPUTERS)
            Add-WordPicture -WordDocument $reportFile -ImagePath $imagePath -Alignment center -ImageWidth 600 -Supress $True
        }

        #ACL
        $computerACL = Get-GPOAclExtended -GPO_ACL $($computer.DistinguishedName)

        Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "$($computer.Name) Permissions Extended" -Supress $true 
        Add-WordTable -WordDocument $reportFile -DataTable $($computerACL | Select-Object -Property * -ExcludeProperty ACLs) -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle "GPO Options" -Transpose -Supress $true
        Add-WordText -WordDocument $reportFile -Text "" -Supress $true
    
        Add-WordTable -WordDocument $reportFile -DataTable $($computerACL.ACLs) -Design MediumShading1Accent5 -AutoFit Window  -Supress $true
        Add-WordText -WordDocument $reportFile -Text "" -Supress $true
}
    
    Add-WordText -WordDocument $reportFile -Text "Computers Table"  -HeadingType Heading2 -Supress $true
    
    Add-WordText -WordDocument $reportFile -Text "Tabela adresacji"  -HeadingType Heading3 -Supress $true
    $table = $($computers | Select-Object DNSHostName, IP4, IP6,Location)
    Add-WordTable -WordDocument $reportFile -DataTable $table -Design ColorfulGridAccent5 -AutoFit Window -Supress $true


    Add-WordText -WordDocument $reportFile -Text "Tabela bezpieczeństwa 1"  -HeadingType Heading3 -Supress $true
    $table = $($computers | Select-Object Name,Enabled,LockedOut,PasswordExpired)
    Add-WordTable -WordDocument $reportFile -DataTable $table -Design ColorfulGridAccent5 -AutoFit Window -Supress $true


    Add-WordText -WordDocument $reportFile -Text "Tabela bezpieczeństwa 2"  -HeadingType Heading3 -Supress $true
    $table = $($computers | Select-Object Name, AllowReversiblePasswordEncryption,CannotChangePassword,PasswordNeverExpires,PasswordNotRequired)
    Add-WordTable -WordDocument $reportFile -DataTable $table -Design ColorfulGridAccent5 -AutoFit Window -Supress $true

    
    Add-WordText -WordDocument $reportFile -Text "Tabela bezpieczeństwa 3"  -HeadingType Heading3 -Supress $true
    $table = $($computers | Select-Object Name, AccountNotDelegated,TrustedForDelegation,IsCriticalSystemObject)
    Add-WordTable -WordDocument $reportFile -DataTable $table -Design ColorfulGridAccent5 -AutoFit Window -Supress $true


    Add-WordText -WordDocument $reportFile -Text "Tabela bezpieczeństwa 4"  -HeadingType Heading3 -Supress $true
    $table = $($computers | Select-Object Name, DoesNotRequirePreAuth,ProtectedFromAccidentalDeletion,USEDESKeyOnly)
    Add-WordTable -WordDocument $reportFile -DataTable $table -Design ColorfulGridAccent5 -AutoFit Window -Supress $true

    Add-WordText -WordDocument $reportFile -Text "Computer charts"  -HeadingType Heading3 -Supress $true
    
    $chart = $computers | Group-Object Enabled | Select-Object Name, @{Name="Values";Expression={$_.Count}}
    Add-WordChart -CType "Barchart" -CData $chart -STitle "Wykresy kont komputerów wyłączonych/włączonych" -CTitle "Stosunek liczby kont komputerów wyłączonych i włączonych"
    
    $chart = $computers | Group-Object OperatingSystem | Select-Object Name, @{Name="Values";Expression={$_.Count}}
    Add-WordChart -CType "Piechart" -CData $chart -STitle "Wykresy stosunku systemów operacyjnych" -CTitle "Stosunek rodzajów systemów operacyjnych"
    
    $chart = $computers | Group-Object OperatingSystemVersion | Select-Object Name, @{Name="Values";Expression={$_.Count}}
    Add-WordChart -CType "Piechart" -CData $chart -STitle "Wykresy stosunku  wersji systemów operacyjnych" -CTitle "Stosunek wersji systemów operacyjnych"
    
    $chart = $computers | Select-Object Name,@{Name="Values";Expression={$_.LogonCount}} | Sort-Object -Descending Values |Select-Object -First 10
    Add-WordChart -CType "Piechart" -CData $chart -STitle "Wykresy najczęściej logujących się komputerów" -CTitle "Wykres najczęściej logujących się komputerów"
    #TODO:LocalPolicyFlags - sprawdzic w pracy

    Add-WordText -WordDocument $reportFile -Text "Computers List"  -HeadingType Heading2 -Supress $true
    
    Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "Ostatnie 10 logujących się komputerów" -Supress $true
    $list = $($($computers | Select-Object LastLogonDate, Name | Sort-Object -Descending LastLogonDate | Select-Object -First 10) | Select-Object @{Name = "ComputerName"; Expression = { "$($_.Name) - $($_.LastLogonDate)" } }).ComputerName
    Add-WordList -WordDocument $reportFile -ListType Numbered -ListData $list -Supress $true -Verbose

    Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "Ostatnie 10 zmienionych haseł komputerów" -Supress $true
    $list = $($($computers | Select-Object PasswordLastSet, Name | Sort-Object -Descending PasswordLastSet | Select-Object -First 10) | Select-Object @{Name = "ComputerName"; Expression = { "$($_.Name) - $($_.PasswordLastSet)" } }).ComputerName
    Add-WordList -WordDocument $reportFile -ListType Numbered -ListData $list -Supress $true -Verbose
    
    Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "Ostatnie 10 zmienionych komputerów" -Supress $true
    $list = $($($computers | Select-Object whenChanged, Name | Sort-Object -Descending whenChanged | Select-Object -First 10) | Select-Object @{Name = "ComputerName"; Expression = { "$($_.Name) - $($_.whenChanged)" } }).ComputerName
    Add-WordList -WordDocument $reportFile -ListType Numbered -ListData $list -Supress $true -Verbose

    Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "Ostatnie 10 utworzonych komputerów" -Supress $true
    $list = $($($computers | Select-Object whenCreated, Name | Sort-Object -Descending whenCreated | Select-Object -First 10) | Select-Object @{Name = "ComputerName"; Expression = { "$($_.Name) - $($_.whenCreated)" } }).ComputerName
    Add-WordList -WordDocument $reportFile -ListType Numbered -ListData $list -Supress $true -Verbose

#endregion COMPUTERS##########################################################################################
##############################################################################################################
Save-WordDocument $reportFile -Supress $true -Language "en-US"  -OpenDocument -Verbose
}
#TODO:Standardy wykonania wykresów i tabel
