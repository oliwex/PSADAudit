function Add-Description 
{
    Param(
        [Parameter(Mandatory = $true)]
        [Alias("DescriptionPath")]
        $pathToDescription,
        [ValidateSet("Organisational Unit","Group","User")]
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
    elseif($descriptionType -like "Group"){

    }elseif($descriptionType -like "User"){

    }

    Add-WordList -WordDocument $reportFile -ListType Bulleted -ListData $descriptionObject -Supress $true -Verbose

}
