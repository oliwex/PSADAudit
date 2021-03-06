function Invoke-ADAudit
{
    
    #Prepare Workplace for script Execution
    $basePath=New-Workplace
    $reportGraphFolders = Get-ReportFolders -BasePath $basePath -GraphFoldersHashtable $graphFolders

    $reportFilePath = Join-Path -Path $basePath -ChildPath "report.docx"
    $logFilePath=Join-Path -Path $basePath -ChildPath "log.txt"

    New-InformationLog -LogPath $logFilePath -Message "Started to creating report" -Color GREEN

    $reportFile = New-WordDocument $reportFilePath

    Add-WordText -WordDocument $reportFile -Text "Active Directory Report" -FontSize 28 -FontFamily 'Calibri Light' -Supress $True
    Add-WordPageBreak -WordDocument $reportFile -Supress $true

    #region TOC #########################################################################################################
    Add-WordTOC -WordDocument $reportFile -Title "Table of Content" -Supress $true

    New-InformationLog -LogPath $logFilePath -Message "Created Table Of Content" -Color GREEN
    
    Add-WordPageBreak -WordDocument $reportFile -Supress $true

    #endregion TOC ########################################################################################################

    #region OU ############################################################################################################
    New-InformationLog -LogPath $logFilePath -Message "Started to create Organizational Unit part" -Color GREEN
    
    Add-WordText -WordDocument $reportFile -HeadingType Heading1 -Text 'Organisational Units List' -Supress $true
    Add-Description -DescriptionPath $pathToDescription -DescriptionType "Organisational Unit"

    $ous = Get-OUInformation
    New-InformationLog -LogPath $logFilePath -Message "Information about Organizational Units aquired" -Color GREEN

    foreach ($ou in $ous) 
    {
        Add-WordText -WordDocument $reportFile -HeadingType Heading2 -Text $($ou.Name) -Supress $true
        
        Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "$($ou.Name) Information" -Supress $true
        Add-WordTable -WordDocument $reportFile -DataTable $ou -Design MediumShading1Accent5 -AutoFit Window -OverwriteTitle $($ou.Name) -Transpose  -Supress $True
        
        New-InformationLog -LogPath $logFilePath -Message "Created table about $($ou.Name) Organizational Unit" -Color GREEN

        Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "$($ou.Name) Graph" -Supress $true 

        $ouLeaf = $(Get-ADOrganizationalUnit -Filter "*" -SearchBase $($ou.DistinguishedName) -SearchScope OneLevel).Name
        $ouRoot=$($($ou.DistinguishedName) -split ',*..=')[2]
        if (($null -eq $ouLeaf) -and ($null -eq $ouRoot))
        {
            Add-WordText -WordDocument $reportFile -Text "$($ou.Name) do not have above and below elements" -Supress $true
            New-InformationLog -LogPath $logFilePath -Message "$($ou.Name) Organizational Unit do not have above and below elements" -Color RED
        }
        else
        { 
            $imagePath = Get-GraphImage -GraphRoot $ouRoot -GraphMiddle $($ou.Name) -GraphLeaf $ouLeaf  -BasePathToGraphImage $($reportGraphFolders.OU)
            Add-WordPicture -WordDocument $reportFile -ImagePath $imagePath -Alignment center -ImageWidth 600 -Supress $True
            
            New-InformationLog -LogPath $logFilePath -Message "$($ou.Name) Organizational Unit have above and below elements" -Color GREEN
        }
        
        #ManagedBy
        #TEST
        Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "$($ou.Name) ManagedBy" -Supress $true 
            
        $ouTMP = $ou.ManagedBy | ForEach-Object { $(($_ -split ',*..=')[1]) }
        if ($null -eq $ouTMP) 
        {
            Add-WordText -WordDocument $reportFile -Text "$($ou.Name) do not have above elements" -Supress $true
            
            New-InformationLog -LogPath $logFilePath -Message "$($ou.Name) Organizational Unit do not have ManagedBy elements" -Color RED       
        }
        else 
        {
            $imagePath = Get-GraphImage -GraphRoot $null -GraphMiddle $ouTMP -GraphLeaf $($ou.Name)  -BasePathToGraphImage $($reportGraphFolders.OU)
            Add-WordPicture -WordDocument $reportFile -ImagePath $imagePath -Alignment center -ImageWidth 600 -Supress $True

            New-InformationLog -LogPath $logFilePath -Message "$($ou.Name) Organizational Unit do have ManagedBy elements" -Color GREEN 
        }

        #ACL
        $ouACL = Get-OUAcl -OU $($ou.DistinguishedName)
        
        Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "$($ou.Name) Permissions" -Supress $true 
        Add-WordTable -WordDocument $reportFile -DataTable $($ouACL | Select-Object -Property * -ExcludeProperty ACLs) -Design MediumShading1Accent5 -AutoFit Window -OverwriteTitle "OU Options" -Transpose -Supress $true #MediumShading1Accent5
        Add-WordText -WordDocument $reportFile -Text "" -Supress $true
        
        New-InformationLog -LogPath $logFilePath -Message "Created $($ou.Name) Organizational Unit ACL Table 1" -Color GREEN 

        Add-WordTable -WordDocument $reportFile -DataTable $($ouACL.ACLs) -Design MediumShading1Accent5 -AutoFit Window  -Supress $true
        Add-WordText -WordDocument $reportFile -Text "" -Supress $true

        New-InformationLog -LogPath $logFilePath -Message "Created $($ou.Name) Organizational Unit ACL Table 2" -Color GREEN 
    }

    Add-WordText -WordDocument $reportFile -Text "Organisational Unit Tables"  -HeadingType Heading2 -Supress $true
    Add-WordText -WordDocument $reportFile -Text "Cities and Country Table"  -HeadingType Heading3 -Supress $true
    $ouTable = $($ous | Select-Object Name, StreetAddress, PostalCode, City, State, Country)
    Add-WordTable -WordDocument $reportFile -DataTable $ouTable -Design ColorfulGridAccent1 -Supress $True #-Verbose

    New-InformationLog -LogPath $logFilePath -Message "Created Organizational Units Table" -Color GREEN 

    Add-WordText -WordDocument $reportFile -Text "Organisational Unit List"  -HeadingType Heading2 -Supress $true

    Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "Last 10 changed organisational unit" -Supress $true
    $list = $($($ous | Select-Object whenChanged, Name | Sort-Object -Descending whenChanged | Select-Object -First 10) | Select-Object @{Name = "OUName"; Expression = { "$($_.Name) - $($_.whenChanged)" } }).OUName
    Add-WordList -WordDocument $reportFile -ListType Numbered -ListData $list -Supress $true -Verbose

    New-InformationLog -LogPath $logFilePath -Message "Created Organizational Units List 1" -Color GREEN 

    Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "Last 10 created organisational unit" -Supress $true
    $list = $($($ous | Select-Object whenCreated, Name | Sort-Object -Descending whenCreated | Select-Object -First 10) | Select-Object @{Name = "OUName"; Expression = { "$($_.Name) - $($_.whenCreated)" } }).OUName
    Add-WordList -WordDocument $reportFile -ListType Numbered -ListData $list -Supress $true -Verbose

    New-InformationLog -LogPath $logFilePath -Message "Created Organizational Units List 2" -Color GREEN 

    #endregion OU #####################################################################################################

    #region GROUPS#####################################################################################################
    New-InformationLog -LogPath $logFilePath -Message "Started to create Group part" -Color GREEN 
    
    Add-WordText -WordDocument $reportFile -Text 'Spis Grup' -HeadingType Heading1 -Supress $true

    Add-Description -DescriptionPath $pathToDescription -DescriptionType "Group" 

    $groups = Get-GROUPInformation

    New-InformationLog -LogPath $logFilePath -Message "Information about groups aquired" -Color GREEN 

    Add-WordText -WordDocument $reportFile -Text "DomainLocal groups"  -HeadingType Heading2 -Supress $true

    $groupObjects=$groups | Where-Object {$_.GroupType -band 1 }
    
    New-InformationLog -LogPath $logFilePath -Message "Started to create domain local group part" -Color GREEN 
    
    foreach ($group in $groupObjects) 
    {
        Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text $($group.Name) -Supress $true
        
        Add-WordText -WordDocument $reportFile -HeadingType Heading4 -Text "$($group.Name) Information" -Supress $true
        Add-WordTable -WordDocument $reportFile -DataTable $group -Design MediumShading1Accent5 -AutoFit Window -OverwriteTitle $($group.Name) -Transpose -Supress $True
        
        New-InformationLog -LogPath $logFilePath -Message "Created table about $($group.Name) Group" -Color GREEN
        
        Add-WordText -WordDocument $reportFile -HeadingType Heading4 -Text "$($group.Name) Graph" -Supress $true


        $groupLeafTMP = $group.Members | ForEach-Object { $(($_ -split ',*..=')[1]) }
        $groupRootTMP = $group.MemberOf | ForEach-Object { $(($_ -split ',*..=')[1]) }
        if (($null -eq $groupLeafTMP) -and ($null -eq $groupRootTMP))
        {
            Add-WordText -WordDocument $reportFile -Text "$($group.Name) do not have above and below elements" -Supress $true
            New-InformationLog -LogPath $logFilePath -Message "$($group.Name) Group do not have above and below elements" -Color RED
        }
        else
        {
            $imagePath = Get-GraphImage -GraphRoot $groupRootTMP -GraphMiddle $($group.Name) -GraphLeaf $groupLeafTMP -pathToImage $($reportGraphFolders.GROUP)
            Add-WordPicture -WordDocument $reportFile -ImagePath $imagePath -Alignment center -ImageWidth 600 -Supress $True
            New-InformationLog -LogPath $logFilePath -Message "$($group.Name) Group have Members or is member of other group" -Color GREEN

        }


        #ManagedBy
        Add-WordText -WordDocument $reportFile -HeadingType Heading4 -Text "$($group.Name) ManagedBy" -Supress $true 
            
        $groupTMP = $group.ManagedBy | ForEach-Object { $(($_ -split ',*..=')[1]) }
        if ($null -eq $groupTMP) 
        {
            Add-WordText -WordDocument $reportFile -Text "$($group.Name) do not have above elements" -Supress $true
            New-InformationLog -LogPath $logFilePath -Message "$($group.Name) Group do not have above elements" -Color RED       
        }
        else 
        {
            $imagePath = Get-GraphImage -GraphRoot $null -GraphMiddle $groupTMP -GraphLeaf $($group.Name)  -BasePathToGraphImage $($reportGraphFolders.GROUP)
            Add-WordPicture -WordDocument $reportFile -ImagePath $imagePath -Alignment center -ImageWidth 600 -Supress $True
            
            New-InformationLog -LogPath $logFilePath -Message "$($group.Name) Group have ManagedBy elements" -Color GREEN
        }


        #ACL
        $groupACL=Get-GROUPAcl -GROUP_ACL $($group.DistinguishedName)

        Add-WordText -WordDocument $reportFile -HeadingType Heading4 -Text "$($group.Name) Permissions" -Supress $true 
        Add-WordTable -WordDocument $reportFile -DataTable $($groupACL | Select-Object -Property * -ExcludeProperty ACLs) -Design MediumShading1Accent5 -AutoFit Window -OverwriteTitle "OU Options" -Transpose -Supress $true
        Add-WordText -WordDocument $reportFile -Text "" -Supress $true
        
        New-InformationLog -LogPath $logFilePath -Message "Created Table about $($group.Name) Group" -Color GREEN

        Add-WordTable -WordDocument $reportFile -DataTable $($groupACL.ACLs) -Design MediumShading1Accent5 -AutoFit Window  -Supress $true
        Add-WordText -WordDocument $reportFile -Text "" -Supress $true

        New-InformationLog -LogPath $logFilePath -Message "Created ACL Table about $($group.Name) Group" -Color GREEN
    }

    Add-WordText -WordDocument $reportFile -Text "Security Groups"  -HeadingType Heading2 -Supress $true

    $groupObjects=$groups | Where-Object { (-not($_.GroupType -band 1)) -and ($_.GroupCategory -eq "Security") }
    
    New-InformationLog -LogPath $logFilePath -Message "Started to create security group part" -Color GREEN

    foreach ($group in $groupObjects) 
    {
        Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text $($group.Name) -Supress $true
        
        Add-WordText -WordDocument $reportFile -HeadingType Heading4 -Text "$($group.Name) Information" -Supress $true
        Add-WordTable -WordDocument $reportFile -DataTable $group -Design MediumShading1Accent5 -AutoFit Window -OverwriteTitle $($group.Name) -Transpose -Supress $True
        
        New-InformationLog -LogPath $logFilePath -Message "Created table about $($group.Name) Group" -Color GREEN

        Add-WordText -WordDocument $reportFile -HeadingType Heading4 -Text "$($group.Name) Graph" -Supress $true

        $groupLeafTMP = $group.Members | ForEach-Object { $(($_ -split ',*..=')[1]) }
        $groupRootTMP = $group.MemberOf | ForEach-Object { $(($_ -split ',*..=')[1]) }
        if (($null -eq $groupLeafTMP) -and ($null -eq $groupRootTMP))
        {
            Add-WordText -WordDocument $reportFile -Text "$($group.Name) do not have above and below elements" -Supress $true
            New-InformationLog -LogPath $logFilePath -Message "$($group.Name) Group do not have above and below elements" -Color RED
        }
        else
        {
            $imagePath = Get-GraphImage -GraphRoot $groupRootTMP -GraphMiddle $($group.Name) -GraphLeaf $groupLeafTMP -pathToImage $($reportGraphFolders.GROUP)
            Add-WordPicture -WordDocument $reportFile -ImagePath $imagePath -Alignment center -ImageWidth 600 -Supress $True
            New-InformationLog -LogPath $logFilePath -Message "$($group.Name) Group have members and is MemberOf other group" -Color GREEN
        }

        #ManagedBy
        Add-WordText -WordDocument $reportFile -HeadingType Heading4 -Text "$($group.Name) ManagedBy" -Supress $true 
            
        $groupTMP = $group.ManagedBy | ForEach-Object { $(($_ -split ',*..=')[1]) }
        if ($null -eq $groupTMP) 
        {
            Add-WordText -WordDocument $reportFile -Text "$($group.Name) do not have above elements" -Supress $true
            New-InformationLog -LogPath $logFilePath -Message "$($group.Name) Group do not have above and below elements" -Color RED       
        }
        else 
        {
            $imagePath = Get-GraphImage -GraphRoot $null -GraphMiddle $groupTMP -GraphLeaf $($group.Name)  -BasePathToGraphImage $($reportGraphFolders.GROUP)
            Add-WordPicture -WordDocument $reportFile -ImagePath $imagePath -Alignment center -ImageWidth 600 -Supress $True
            New-InformationLog -LogPath $logFilePath -Message "$($group.Name) Group have ManagedBy elements" -Color GREEN
        }

        #ACL
        $groupACL=Get-GROUPAcl -GROUP_ACL $($group.DistinguishedName)

        Add-WordText -WordDocument $reportFile -HeadingType Heading4 -Text "$($group.Name) Permissions" -Supress $true 
        Add-WordTable -WordDocument $reportFile -DataTable $($groupACL | Select-Object -Property * -ExcludeProperty ACLs) -Design MediumShading1Accent5 -AutoFit Window -OverwriteTitle "OU Options" -Transpose -Supress $true
        Add-WordText -WordDocument $reportFile -Text "" -Supress $true
        
        New-InformationLog -LogPath $logFilePath -Message "Created Table about $($group.Name) Group" -Color GREEN

        Add-WordTable -WordDocument $reportFile -DataTable $($groupACL.ACLs) -Design MediumShading1Accent5 -AutoFit Window  -Supress $true
        Add-WordText -WordDocument $reportFile -Text "" -Supress $true

        New-InformationLog -LogPath $logFilePath -Message "Created ACL Table about $($group.Name) Group" -Color GREEN
    }

    Add-WordText -WordDocument $reportFile -Text "Distribution Groups"  -HeadingType Heading2 -Supress $true

    $groupObjects=$groups | Where-Object {$_.GroupCategory -eq "Distribution" }
    
    New-InformationLog -LogPath $logFilePath -Message "Started to create distribution group part" -Color GREEN

    foreach ($group in $groupObjects) 
    {
        Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text $($group.Name) -Supress $true
        
        Add-WordText -WordDocument $reportFile -HeadingType Heading4 -Text "$($group.Name) Information" -Supress $true
        Add-WordTable -WordDocument $reportFile -DataTable $group -Design MediumShading1Accent5 -AutoFit Window -OverwriteTitle $($group.Name) -Transpose -Supress $True
        
        New-InformationLog -LogPath $logFilePath -Message "Created table about $($group.Name) Group" -Color GREEN

        Add-WordText -WordDocument $reportFile -HeadingType Heading4 -Text "$($group.Name) Graph" -Supress $true

        $groupLeafTMP = $group.Members | ForEach-Object { $(($_ -split ',*..=')[1]) }
        $groupRootTMP = $group.MemberOf | ForEach-Object { $(($_ -split ',*..=')[1]) }
        if (($null -eq $groupLeafTMP) -and ($null -eq $groupRootTMP))
        {
            Add-WordText -WordDocument $reportFile -Text "$($group.Name) do not have above and below elements" -Supress $true
            New-InformationLog -LogPath $logFilePath -Message "$($group.Name) Group do not have above and below elements" -Color RED
        }
        else
        {
            $imagePath = Get-GraphImage -GraphRoot $groupRootTMP -GraphMiddle $($group.Name) -GraphLeaf $groupLeafTMP -pathToImage $($reportGraphFolders.GROUP)
            Add-WordPicture -WordDocument $reportFile -ImagePath $imagePath -Alignment center -ImageWidth 600 -Supress $True
            New-InformationLog -LogPath $logFilePath -Message "$($group.Name) Group have members and is MemberOf other group" -Color GREEN
        }

        #ManagedBy
        Add-WordText -WordDocument $reportFile -HeadingType Heading4 -Text "$($group.Name) ManagedBy" -Supress $true 
            
        $groupTMP = $group.ManagedBy | ForEach-Object { $(($_ -split ',*..=')[1]) }
        if ($null -eq $groupTMP) 
        {
            Add-WordText -WordDocument $reportFile -Text "$($group.Name) do not have above elements" -Supress $true
            New-InformationLog -LogPath $logFilePath -Message "$($group.Name) Group do not have above and below elements" -Color RED         
        }
        else 
        {
            $imagePath = Get-GraphImage -GraphRoot $null -GraphMiddle $groupTMP -GraphLeaf $($group.Name)  -BasePathToGraphImage $($reportGraphFolders.GROUP)
            Add-WordPicture -WordDocument $reportFile -ImagePath $imagePath -Alignment center -ImageWidth 600 -Supress $True
            New-InformationLog -LogPath $logFilePath -Message "$($group.Name) Group have ManagedBy elements" -Color GREEN
        }

        #ACL
        $groupACL=Get-GROUPAcl -GROUP_ACL $($group.DistinguishedName)

        Add-WordText -WordDocument $reportFile -HeadingType Heading4 -Text "$($group.Name) Permissions" -Supress $true 
        Add-WordTable -WordDocument $reportFile -DataTable $($groupACL | Select-Object -Property * -ExcludeProperty ACLs) -Design MediumShading1Accent5 -AutoFit Window -OverwriteTitle "OU Options" -Transpose -Supress $true
        Add-WordText -WordDocument $reportFile -Text "" -Supress $true

        New-InformationLog -LogPath $logFilePath -Message "Created Table about $($group.Name) Group" -Color GREEN
        
        Add-WordTable -WordDocument $reportFile -DataTable $($groupACL.ACLs) -Design MediumShading1Accent5 -AutoFit Window  -Supress $true
        Add-WordText -WordDocument $reportFile -Text "" -Supress $true

        New-InformationLog -LogPath $logFilePath -Message "Created ACL Table about $($group.Name) Group" -Color GREEN
    }

    Add-WordText -WordDocument $reportFile -Text "Group Tables"  -HeadingType Heading2 -Supress $true
    Add-WordText -WordDocument $reportFile -Text "Grup difference"  -HeadingType Heading3 -Supress $true

    $groupTable = $groups | Group-Object GroupScope | ForEach-Object {
        $categories = $_.Group | Group-Object GroupCategory -AsHashtable -AsString

        [PSCustomObject]@{
            GroupName    = $_.Name
            Security     = $categories['Security'].Count
            Distribution = $categories['Distribution'].Count
        }
    }

    Add-WordTable -WordDocument $reportFile -DataTable $groupTable -Design ColorfulGridAccent1 -Supress $True #-Verbose
    
    New-InformationLog -LogPath $logFilePath -Message "Created Group Tables" -Color GREEN

    Add-WordText -WordDocument $reportFile -Text "Group Lists"  -HeadingType Heading2 -Supress $true

    $list = $($($groups | Select-Object whenChanged, Name | Sort-Object -Descending whenChanged | Select-Object -First 10) | Select-Object @{Name = "GroupName"; Expression = { "$($_.Name) - $($_.whenChanged)" } }).GroupName
    Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "Last 10 changed groups" -Supress $true
    Add-WordList -WordDocument $reportFile -ListType Numbered -ListData $list -Supress $true -Verbose

    $list = $($($groups | Select-Object whenCreated, Name | Sort-Object -Descending whenCreated | Select-Object -First 10) | Select-Object @{Name = "GroupName"; Expression = { "$($_.Name) - $($_.whenCreated)" } }).GroupName
    Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "Last 10 created groups" -Supress $true
    Add-WordList -WordDocument $reportFile -ListType Numbered -ListData $list -Supress $true -Verbose

    $list = $($groups | Where-Object { $_.Members.Count -eq 0 }).Name
    Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "Empty Groups" -Supress $true
    Add-WordList -WordDocument $reportFile -ListType Numbered -ListData $list -Supress $true -Verbose

    New-InformationLog -LogPath $logFilePath -Message "Created Group Lists" -Color GREEN

    Add-WordText -WordDocument $reportFile -Text "Group Charts"  -HeadingType Heading2 -Supress $true

    $chart = $groups | Group-Object GroupCategory | Select-Object Name, @{Name="Values";Expression={$_.Count}}
    Add-WordChart -CType "Barchart" -CData $chart -STitle "Distribution\Security group chart" -CTitle "Ratio between security groups and distribution groups"

    $chart = $groups | Group-Object GroupScope | Select-Object Name, @{Name="Values";Expression={$_.Count}}
    Add-WordChart -CType "Barchart" -CData $chart -STitle "Local\Global\Universal group chart" -CTitle "Ratio between local groups,global groups and universal groups"

    New-InformationLog -LogPath $logFilePath -Message "Created Group Graphs" -Color GREEN

    #endregion GROUPS#####################################################################################################

    #region USERS#####################################################################################################
    New-InformationLog -LogPath $logFilePath -Message "Started to create user part" -Color GREEN
    
    Add-WordText -WordDocument $reportFile -Text 'Users List' -HeadingType Heading1 -Supress $true
    Add-Description -DescriptionPath $pathToDescription -DescriptionType "User"

    $users = Get-USERInformation

    New-InformationLog -LogPath $logFilePath -Message "Information about users aquired" -Color GREEN

    foreach ($user in $users) 
    {
        Add-WordText -WordDocument $reportFile -HeadingType Heading2 -Text $($user.Name) -Supress $true
        
        Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "$($user.Name) Information" -Supress $true
        Add-WordTable -WordDocument $reportFile -DataTable $user -Design MediumShading1Accent5 -AutoFit Window -OverwriteTitle $($user.Name) -Transpose -Supress $true
    
        New-InformationLog -LogPath $logFilePath -Message "Created table about $($user.Name)" -Color GREEN
        
        #MemberOf
        Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "$($user.Name) MemberOfGroup Graph" -Supress $true 
        
        if ($null -eq $($user.MemberOf)) {
            Add-WordText -WordDocument $reportFile -Text "$($user.Name) do not have below elements" -Supress $true
            $memberOfTMP = $($($($user.PrimaryGroup) -split ',*..=')[1])
            
            $imagePath = Get-GraphImage -GraphRoot $null -GraphMiddle $($user.Name) -GraphLeaf $memberOfTMP  -BasePathToGraphImage $($reportGraphFolders.USERS)
            Add-WordPicture -WordDocument $reportFile -ImagePath $imagePath -Alignment center -ImageWidth 600 -Supress $True

            New-InformationLog -LogPath $logFilePath -Message "User $($user.Name) is member of primary group" -Color RED
        }
        else {
            $memberOfTMP = $($($user.MemberOf) + $($user.PrimaryGroup) | ForEach-Object { $(($_ -split ',*..=')[1]) }  )

            $imagePath = Get-GraphImage -GraphRoot $null -GraphMiddle $($user.Name) -GraphLeaf $memberOfTMP  -BasePathToGraphImage $($reportGraphFolders.USERS)
            Add-WordPicture -WordDocument $reportFile -ImagePath $imagePath -Alignment center -ImageWidth 600 -Supress $True

            New-InformationLog -LogPath $logFilePath -Message "User $($user.Name) is member of primary group and other group" -Color RED
        }

        #ManagedBy
        Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "$($user.Name) ManagedBy" -Supress $true 
            
        $userTMP = $user.ManagedBy | ForEach-Object { $(($_ -split ',*..=')[1]) }
        if ($null -eq $userTMP) 
        {
            Add-WordText -WordDocument $reportFile -Text "$($user.Name) do not have above elements" -Supress $true
            New-InformationLog -LogPath $logFilePath -Message " User $($user.Name) do not have above elements" -Color RED      
        }
        else 
        {
            $imagePath = Get-GraphImage -GraphRoot $null -GraphMiddle $userTMP -GraphLeaf $($user.Name)  -BasePathToGraphImage $($reportGraphFolders.USERS)
            Add-WordPicture -WordDocument $reportFile -ImagePath $imagePath -Alignment center -ImageWidth 600 -Supress $True
            New-InformationLog -LogPath $logFilePath -Message "User $($user.Name) have ManagedBy elements" -Color GREEN
        }

        #Manager
        Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "$($user.Name) DirectManager" -Supress $true 
        $managerTMP = $user.Manager| ForEach-Object { $(($_ -split ',*..=')[1]) } 
        $directReportsTMP = $user.DirectReports | ForEach-Object { $(($_ -split ',*..=')[1]) }
        if (($null -eq $managerTMP) -and ($null -eq $directReportsTMP))
        {
            Add-WordText -WordDocument $reportFile -Text "$($user.Name) do not have above and below elements" -Supress $true
            New-InformationLog -LogPath $logFilePath -Message "User $($user.Name) do not have above and below elements" -Color RED   
        }
        else
        {
            $imagePath = Get-GraphImage -GraphRoot $managerTMP -GraphMiddle $($user.Name) -GraphLeaf $directReportsTMP  -BasePathToGraphImage $($reportGraphFolders.USERS)
            Add-WordPicture -WordDocument $reportFile -ImagePath $imagePath -Alignment center -ImageWidth 600 -Supress $True
            New-InformationLog -LogPath $logFilePath -Message "User $($user.Name) have Manager or Direct Employees" -Color GREEN
        }
        #TODO:Create graph with full organisation manager and direct report

        #ACL
        $userACL = Get-USERAcl -USER_ACL $($user.DistinguishedName)

        Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "$($user.Name) Permissions" -Supress $true 
        Add-WordTable -WordDocument $reportFile -DataTable $($userACL | Select-Object -Property * -ExcludeProperty ACLs) -Design MediumShading1Accent5 -AutoFit Window -OverwriteTitle "User Options" -Transpose -Supress $true
        Add-WordText -WordDocument $reportFile -Text "" -Supress $true
        
        New-InformationLog -LogPath $logFilePath -Message "Created Table about $($user.Name) User" -Color GREEN

        Add-WordTable -WordDocument $reportFile -DataTable $($userACL.ACLs) -Design MediumShading1Accent5 -AutoFit Window  -Supress $true
        Add-WordText -WordDocument $reportFile -Text "" -Supress $true

        New-InformationLog -LogPath $logFilePath -Message "Created ACL Table about $($user.Name) User" -Color GREEN

    }

    Add-WordText -WordDocument $reportFile -Text "Users Table"  -HeadingType Heading2 -Supress $true

    Add-WordText -WordDocument $reportFile -Text "Users Table Location"  -HeadingType Heading3 -Supress $true
    $table = $($users | Select-Object Name, Department, City, Country)
    Add-WordTable -WordDocument $reportFile -DataTable $table -Design MediumShading1Accent5 -AutoFit Window -Supress $true

    Add-WordText -WordDocument $reportFile -Text "Security Table"  -HeadingType Heading3 -Supress $true
    $table = $($users | Select-Object Name, CannotChangePassword, PasswordExpired, PasswordNeverExpires, PasswordNotRequired)
    Add-WordTable -WordDocument $reportFile -DataTable $table -Design MediumShading1Accent5 -AutoFit Window -Supress $true

    New-InformationLog -LogPath $logFilePath -Message "Created Tables about Users" -Color GREEN

    Add-WordText -WordDocument $reportFile -Text "Users Lists"  -HeadingType Heading2 -Supress $true

    Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "Last 10 changed users" -Supress $true
    $list = $($($users | Select-Object whenChanged, Name | Sort-Object -Descending whenChanged | Select-Object -First 10) | Select-Object @{Name = "UserName"; Expression = { "$($_.Name) - $($_.whenChanged)" } }).UserName
    Add-WordList -WordDocument $reportFile -ListType Numbered -ListData $list -Supress $true -Verbose

    Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "Last 10 created users" -Supress $true
    $list = $($($users | Select-Object whenCreated, Name | Sort-Object -Descending whenCreated | Select-Object -First 10) | Select-Object @{Name = "UserName"; Expression = { "$($_.Name) - $($_.whenCreated)" } }).UserName
    Add-WordList -WordDocument $reportFile -ListType Numbered -ListData $list -Supress $true -Verbose

    New-InformationLog -LogPath $logFilePath -Message "Created Lists about Users" -Color GREEN

    Add-WordText -WordDocument $reportFile -Text "User Charts"  -HeadingType Heading2 -Supress $true

    $chart = $users | Group-Object Enabled | Select-Object Name, @{Name="Values";Expression={$_.Count}}
        Add-WordChart -CType "Barchart" -CData $chart -STitle "Enabled\Disabled accounts chart" -CTitle "The ratio of the number of disabled and enabled accounts"

    $chart = $users | Group-Object Office | Select-Object Name , @{Name = "Values"; Expression = {[math]::round(((($_.Count) / $users.Count) * 100), 2)} } | Where-Object { $_.Values -ge 1 } | Sort-Object -Descending Values
    Add-WordChart -CType "Piechart" -CData $chart -STitle "Chart of offices in a cross section of the company" -CTitle "ratio of the number of positions"

    $chart = $users | Group-Object Title | Select-Object Name, @{Name = "Values"; Expression = { [math]::round(((($_.Count) / $users.Count) * 100), 2)} } | Where-Object { $_.Values -ge 1 } | Sort-Object -Descending Values
        Add-WordChart -CType "Piechart" -CData $chart -STitle "Chart of positions in the cross-section of the companyy" -CTitle "Chart of positions in the cross-section of the company"

    $chart = $users | Group-Object Department | Select-Object Name, @{Name = "Values"; Expression = {[math]::round(((($_.Count) / $users.Count) * 100), 2)} } | Where-Object { $_.Values -ge 1 } | Sort-Object -Descending Values
    Add-WordChart -CType "Piechart" -CData $chart -STitle "Chart of departments in a cross section of the company" -CTitle "Chart of positions in the cross-section of the company"

    New-InformationLog -LogPath $logFilePath -Message "Created Charts about Users" -Color GREEN

    #endregion USERS#####################################################################################################

    #region GPO############################################################################################################
    New-InformationLog -LogPath $logFilePath -Message "Started to create GPO part" -Color GREEN
    
    Add-WordText -WordDocument $reportFile -HeadingType Heading1 -Text 'Group Policy Lists' -Supress $true
    Add-Description -DescriptionPath $pathToDescription -DescriptionType "GPOPolicy"

    $groupPolicyObjects = Get-GPO -Domain $($Env:USERDNSDOMAIN) -All 

    New-InformationLog -LogPath $logFilePath -Message "Information about GPO aquired" -Color GREEN

    $groupPolicyObjectsList = foreach ($groupPolicyObject in $groupPolicyObjects) 
    {
        $gpoObject = Get-GPOPolicy -GroupPolicy $groupPolicyObject
        
        Add-WordText -WordDocument $reportFile -HeadingType Heading2 -Text $($gpoObject.Name) -Supress $true
        
        Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "$($gpoObject.Name) Information" -Supress $true
        Add-WordTable -WordDocument $reportFile -DataTable $gpoObject -Design MediumShading1Accent5 -AutoFit Window -OverwriteTitle $($gpoObject.Name) -Transpose -Supress $true
        
        New-InformationLog -LogPath $logFilePath -Message "Created table about $($gpoObject.Name)" -Color GREEN
        
        Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "$($gpoObject.Name) Graph" -Supress $true
        
        if ($null -eq $($gpoObject.Links)) 
        {
            Add-WordText -WordDocument $reportFile -Text "$($gpoObject.Name) do not have below elements" -Supress $true
            New-InformationLog -LogPath $logFilePath -Message "GPO $($gpoObject.Name) do not have links" -Color RED      
        }
        else 
        {
            $linksTMP = $gpoObject.Links.split(";")
            $imagePath = Get-GraphImage -GraphRoot $null -GraphMiddle $($gpoObject.Name) -GraphLeaf $linksTMP -pathToImage $($reportGraphFolders.GPO)
            Add-WordPicture -WordDocument $reportFile -ImagePath $imagePath -Alignment center -ImageWidth 600 -Supress $True

            New-InformationLog -LogPath $logFilePath -Message "GPO $($gpoObject.Name) have links into few AD objects" -Color GREEN 
        }
        
        #ACL
        Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "$($gpoObject.Name) Permissions Simple" -Supress $true
        $gpoObjectACL = Get-GPOAclSimple -GroupPolicy $groupPolicyObject
        
        $gpoObjectACL.ACL | ForEach-Object {
            Add-WordTable -WordDocument $reportFile -DataTable $($_) -Design MediumShading1Accent5 -AutoFit Window  -Supress $true 
        }
        
        New-InformationLog -LogPath $logFilePath -Message "Create ACL tables with GPO $($gpoObject.Name)" -Color GREEN 

        $pathACL = Get-GPOAclExtended -GPO_ACL $($groupPolicyObject.Path)

        Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "$($gpoObject.Name) Permissions Extended" -Supress $true 
        Add-WordTable -WordDocument $reportFile -DataTable $($pathACL | Select-Object -Property * -ExcludeProperty ACLs) -Design MediumShading1Accent5 -AutoFit Window -OverwriteTitle "GPO Options" -Transpose -Supress $true
        Add-WordText -WordDocument $reportFile -Text "" -Supress $true
        
        New-InformationLog -LogPath $logFilePath -Message "Create table with GPO $($gpoObject.Name) information" -Color GREEN 

        Add-WordTable -WordDocument $reportFile -DataTable $($pathACL.ACLs) -Design MediumShading1Accent5 -AutoFit Window  -Supress $true
        Add-WordText -WordDocument $reportFile -Text "" -Supress $true
                
        New-InformationLog -LogPath $logFilePath -Message "Create ACL table with GPO $($gpoObject.Name) information" -Color GREEN 

        $gpoObject
    }


    Add-WordText -WordDocument $reportFile -Text "GroupPolicy Tables"  -HeadingType Heading2 -Supress $true

    Add-WordText -WordDocument $reportFile -Text "Group Policy table 1"  -HeadingType Heading3 -Supress $true
    $gpoTable = $($groupPolicyObjectsList | Select-Object Name, HasComputerSettings, HasUserSettings, ComputerSettings, UserSettings)
    Add-WordTable -WordDocument $reportFile -DataTable $gpoTable -Design ColorfulGridAccent1 -Supress $True #-Verbose

    Add-WordText -WordDocument $reportFile -Text "Group Policy table 2"  -HeadingType Heading3 -Supress $true
    $gpoTable = $($groupPolicyObjectsList | Select-Object Name,UserEnabled, ComputerEnabled)
    Add-WordTable -WordDocument $reportFile -DataTable $gpoTable -Design ColorfulGridAccent1 -Supress $True #-Verbose

    New-InformationLog -LogPath $logFilePath -Message "Created tables with information about GPO" -Color GREEN 

    Add-WordText -WordDocument $reportFile -Text "Group Policy Lists"  -HeadingType Heading2 -Supress $true 

    Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "Last 10 modified organisational unit" -Supress $true
    $list = $($($groupPolicyObjectsList | Select-Object ModificationTime, Name | Sort-Object -Descending ModificationTime | Select-Object -First 10) | Select-Object @{Name = "GPOName"; Expression = { "$($_.Name) - $($_.ModificationTime)" } }).GPOName
    Add-WordList -WordDocument $reportFile -ListType Numbered -ListData $list -Supress $true -Verbose

    Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "Last 10 created organisational unit" -Supress $true
    $list = $($($groupPolicyObjectsList | Select-Object CreationTime, Name | Sort-Object -Descending CreationTime | Select-Object -First 10) | Select-Object @{Name = "GPOName"; Expression = { "$($_.Name) - $($_.CreationTime)" } }).GPOName
    Add-WordList -WordDocument $reportFile -ListType Numbered -ListData $list -Supress $true -Verbose

    Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "Not applied group policies" -Supress $true
    $list = $($groupPolicyObjectsList | Where-Object { $_.Links.Count -eq 0 }).Name
    Add-WordList -WordDocument $reportFile -ListType Numbered -ListData $list -Supress $true -Verbose
        
    New-InformationLog -LogPath $logFilePath -Message "Created Lists with information about GPO" -Color GREEN 

    #endregion GPO################################################################################################

    #region FGPP##################################################################################################
    New-InformationLog -LogPath $logFilePath -Message "Started to create FGPP part" -Color GREEN
    
    Add-WordText -WordDocument $reportFile -HeadingType Heading1 -Text 'Fine Grained Password Policies List' -Supress $true
    Add-Description -DescriptionPath $pathToDescription -DescriptionType "FineGrainedPasswordPolicy"

    $fgpps = Get-FineGrainedPolicies

    New-InformationLog -LogPath $logFilePath -Message "Information about GPO aquired" -Color GREEN

    foreach ($fgpp in $fgpps) 
    {
        Add-WordText -WordDocument $reportFile -HeadingType Heading2 -Text $($fgpp.Name) -Supress $true
        
        Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "$($fgpp.Name) Information" -Supress $true
        Add-WordTable -WordDocument $reportFile -DataTable $fgpp -Design MediumShading1Accent5 -AutoFit Window -OverwriteTitle $($fgpp.Name) -Transpose -Supress $true
                
        New-InformationLog -LogPath $logFilePath -Message "Created table about $($fgpp.Name)" -Color GREEN
        
        Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "$($fgpp.Name) is applied to" -Supress $true
        
        if ($null -eq $($fgpp.'Applies To')) 
        {
            Add-WordText -WordDocument $reportFile -Text "$($fgpp.Name) do not have below elements" -Supress $true
            New-InformationLog -LogPath $logFilePath -Message "FGPP is not applied to objects" -Color RED
        }
        else 
        {
            $fgppAplliedTMP = $($fgpp.'Applies To').split(";")
            $imagePath = Get-GraphImage -GraphRoot $null -GraphMiddle $($fgpp.Name) -GraphLeaf $fgppAplliedTMP -pathToImage $reportGraphFolders.FGPP
            Add-WordPicture -WordDocument $reportFile -ImagePath $imagePath -Alignment center -ImageWidth 600 -Supress $True
            
            New-InformationLog -LogPath $logFilePath -Message "FGPP $($fgpp.Name) is applied to few objects" -Color GREEN
        }

    }
    #endregion FGPP###############################################################################################
    #region COMPUTERS#############################################################################################
    New-InformationLog -LogPath $logFilePath -Message "Started to create Computer part" -Color GREEN
    
    Add-WordText -WordDocument $reportFile -HeadingType Heading1 -Text 'Computers List' -Supress $true
    Add-Description -DescriptionPath $pathToDescription -DescriptionType "Computer"

    $computers=Get-ComputerInformation

    New-InformationLog -LogPath $logFilePath -Message "Information about GPO aquired" -Color GREEN
    foreach ($computer in $computers)
    {
            Add-WordText -WordDocument $reportFile -HeadingType Heading2 -Text $($computer.Name) -Supress $true
        
            Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "$($computer.Name) Information" -Supress $true
            Add-WordTable -WordDocument $reportFile -DataTable $computer -Design MediumShading1Accent5 -AutoFit Window -OverwriteTitle $($computer.Name) -Transpose -Supress $true
                    
            New-InformationLog -LogPath $logFilePath -Message "Created table about $($fgpp.Name)" -Color GREEN
            
            #MemberOf
            Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "$($computer.Name) MemberOfGroup Graph" -Supress $true 
            
            if ($null -eq $computerLeafTMP) 
            {
                Add-WordText -WordDocument $reportFile -Text "$($computer.Name) do not have below elements" -Supress $true
            
                $computerLeafTMP = $($($($computer.PrimaryGroup) -split ',*..=')[1])
                $imagePath = Get-GraphImage -GraphRoot $null -GraphMiddle $($computer.Name) -GraphLeaf $computerLeafTMP  -BasePathToGraphImage $($reportGraphFolders.COMPUTERS)
                Add-WordPicture -WordDocument $reportFile -ImagePath $imagePath -Alignment center -ImageWidth 600 -Supress $True
                
                New-InformationLog -LogPath $logFilePath -Message "Computer $($computer.Name) is not Member Of groups" -Color RED 
            }
            else 
            {        
                $computerLeafTMP = $($($computer.PrimaryGroup) + $($computer.MemberOf) | ForEach-Object { $(($_ -split ',*..=')[1]) }  )
                $imagePath = Get-GraphImage -GraphRoot $null -GraphMiddle $($computer.Name) -GraphLeaf $computerLeafTMP  -BasePathToGraphImage $($reportGraphFolders.COMPUTERS)
                Add-WordPicture -WordDocument $reportFile -ImagePath $imagePath -Alignment center -ImageWidth 600 -Supress $True

                New-InformationLog -LogPath $logFilePath -Message "Computer $($computer.Name) is member of few groups" -Color GREEN
            }

            #ManagedBy
            Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "$($computer.Name) ManagedBy" -Supress $true 
            
            $managerTMP = $computer.ManagedBy | ForEach-Object { $(($_ -split ',*..=')[1]) }
            if ($null -eq $managerTMP) 
            {
                Add-WordText -WordDocument $reportFile -Text "$($computer.Name) do not have above elements" -Supress $true
                New-InformationLog -LogPath $logFilePath -Message "Computer $($computer.Name) is not Managed by other elements" -Color RED     
            }
            else 
            {
                $imagePath = Get-GraphImage -GraphRoot $null -GraphMiddle $managerTMP -GraphLeaf $($computer.Name)  -BasePathToGraphImage $($reportGraphFolders.COMPUTERS)
                Add-WordPicture -WordDocument $reportFile -ImagePath $imagePath -Alignment center -ImageWidth 600 -Supress $True

                New-InformationLog -LogPath $logFilePath -Message "Computer $($computer.Name) is Managed by one of elements" -Color GREEN 
            }

            #ACL
            $computerACL = Get-GPOAclExtended -GPO_ACL $($computer.DistinguishedName)

            Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "$($computer.Name) Permissions Extended" -Supress $true 
            Add-WordTable -WordDocument $reportFile -DataTable $($computerACL | Select-Object -Property * -ExcludeProperty ACLs) -Design MediumShading1Accent5 -AutoFit Window -OverwriteTitle "GPO Options" -Transpose -Supress $true
            Add-WordText -WordDocument $reportFile -Text "" -Supress $true
        
            New-InformationLog -LogPath $logFilePath -Message "Create table with Computer $($computer.Name) information" -Color GREEN 

            Add-WordTable -WordDocument $reportFile -DataTable $($computerACL.ACLs) -Design MediumShading1Accent5 -AutoFit Window  -Supress $true
            Add-WordText -WordDocument $reportFile -Text "" -Supress $true

            New-InformationLog -LogPath $logFilePath -Message "Create ACL table with Computer $($computer.Name)" -Color GREEN 
    }
        
        Add-WordText -WordDocument $reportFile -Text "Computers Table"  -HeadingType Heading2 -Supress $true
        
        Add-WordText -WordDocument $reportFile -Text "Address Table"  -HeadingType Heading3 -Supress $true
        $table = $($computers | Select-Object DNSHostName, IP4, IP6,Location)
        Add-WordTable -WordDocument $reportFile -DataTable $table -Design MediumShading1Accent5 -AutoFit Window -Supress $true


        Add-WordText -WordDocument $reportFile -Text "Security Table 1"  -HeadingType Heading3 -Supress $true
        $table = $($computers | Select-Object Name,Enabled,LockedOut,PasswordExpired)
        Add-WordTable -WordDocument $reportFile -DataTable $table -Design MediumShading1Accent5 -AutoFit Window -Supress $true


        Add-WordText -WordDocument $reportFile -Text "Security Table 2"  -HeadingType Heading3 -Supress $true
        $table = $($computers | Select-Object Name, AllowReversiblePasswordEncryption,CannotChangePassword,PasswordNeverExpires,PasswordNotRequired)
        Add-WordTable -WordDocument $reportFile -DataTable $table -Design MediumShading1Accent5 -AutoFit Window -Supress $true

        
        Add-WordText -WordDocument $reportFile -Text "Security Table 3"  -HeadingType Heading3 -Supress $true
        $table = $($computers | Select-Object Name, AccountNotDelegated,TrustedForDelegation,IsCriticalSystemObject)
        Add-WordTable -WordDocument $reportFile -DataTable $table -Design MediumShading1Accent5 -AutoFit Window -Supress $true


        Add-WordText -WordDocument $reportFile -Text "Security Table 4"  -HeadingType Heading3 -Supress $true
        $table = $($computers | Select-Object Name, DoesNotRequirePreAuth,ProtectedFromAccidentalDeletion,USEDESKeyOnly)
        Add-WordTable -WordDocument $reportFile -DataTable $table -Design MediumShading1Accent5 -AutoFit Window -Supress $true

        New-InformationLog -LogPath $logFilePath -Message "Create tables with Computers information" -Color GREEN 

        Add-WordText -WordDocument $reportFile -Text "Computer charts"  -HeadingType Heading3 -Supress $true
        
        $chart = $computers | Group-Object Enabled | Select-Object Name, @{Name="Values";Expression={$_.Count}}
        Add-WordChart -CType "Barchart" -CData $chart -STitle "Computer account enabled\disabled chart" -CTitle "The ratio of the number of computer accounts enabled and disabled"
        
        $chart = $computers | Group-Object OperatingSystem | Select-Object Name, @{Name="Values";Expression={$_.Count}}
        Add-WordChart -CType "Piechart" -CData $chart -STitle "Operating systems ratio charts" -CTitle "The ratio of the types of operating systems"
        
        $chart = $computers | Group-Object OperatingSystemVersion | Select-Object Name, @{Name="Values";Expression={$_.Count}}
        Add-WordChart -CType "Piechart" -CData $chart -STitle "Operating system version ratio charts" -CTitle "The ratio of the versions of the operating systems"
        
        $chart = $computers | Select-Object Name,@{Name="Values";Expression={$_.LogonCount}} | Sort-Object -Descending Values |Select-Object -First 10
        Add-WordChart -CType "Piechart" -CData $chart -STitle "Charts of the most frequently logged in computers" -CTitle "The chart of the most frequently logged on computers"

        New-InformationLog -LogPath $logFilePath -Message "Create charts with Computers information" -Color GREEN 

        Add-WordText -WordDocument $reportFile -Text "Computers List"  -HeadingType Heading2 -Supress $true
        
        Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "Last 10 logging in computers" -Supress $true
        $list = $($($computers | Select-Object LastLogonDate, Name | Sort-Object -Descending LastLogonDate | Select-Object -First 10) | Select-Object @{Name = "ComputerName"; Expression = { "$($_.Name) - $($_.LastLogonDate)" } }).ComputerName
        Add-WordList -WordDocument $reportFile -ListType Numbered -ListData $list -Supress $true -Verbose

        Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "Last 10 computer passwords changed" -Supress $true
        $list = $($($computers | Select-Object PasswordLastSet, Name | Sort-Object -Descending PasswordLastSet | Select-Object -First 10) | Select-Object @{Name = "ComputerName"; Expression = { "$($_.Name) - $($_.PasswordLastSet)" } }).ComputerName
        Add-WordList -WordDocument $reportFile -ListType Numbered -ListData $list -Supress $true -Verbose
        
        Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "Last 10 changed computers" -Supress $true
        $list = $($($computers | Select-Object whenChanged, Name | Sort-Object -Descending whenChanged | Select-Object -First 10) | Select-Object @{Name = "ComputerName"; Expression = { "$($_.Name) - $($_.whenChanged)" } }).ComputerName
        Add-WordList -WordDocument $reportFile -ListType Numbered -ListData $list -Supress $true -Verbose

        Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text "Last 10 computers created" -Supress $true
        $list = $($($computers | Select-Object whenCreated, Name | Sort-Object -Descending whenCreated | Select-Object -First 10) | Select-Object @{Name = "ComputerName"; Expression = { "$($_.Name) - $($_.whenCreated)" } }).ComputerName
        Add-WordList -WordDocument $reportFile -ListType Numbered -ListData $list -Supress $true -Verbose

        New-InformationLog -LogPath $logFilePath -Message "Create Lists with Computers information" -Color GREEN 
    #endregion COMPUTERS##########################################################################################
    ##############################################################################################################
    Save-WordDocument $reportFile -Supress $true -Language "en-US" -Verbose
}