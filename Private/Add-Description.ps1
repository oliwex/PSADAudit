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
