# PSADAudit
Information about PSADAudit
[TODO:images]
# Introduction
The purpose of this project is provide information about OrganizationalUnits, Users, Groups, Computers GPO in Word document in case of organisational Audit. The module is a simple script which uses multiple functions to reach his goal. The scipt shows users in graphs, tables and images to better visualise his data.
# Technologies
* PowerShell Modules https://github.com/KevinMarquette/PSGraph
    * PSWriteWord - Author: [EvotecIT - Przemyslaw Klys](https://github.com/EvotecIT/PSWriteWord)
    * PSGraph - Author: [Kevin Marquette](https://github.com/KevinMarquette/PSGraph)
* Graphviz - simple API to create graphs in PowerShell - [Docs](https://graphviz.org/)
# Requirements
* PowerShell Min Version 5.1
* Graphviz
# Functions
## Private
### Add-Description
    <details>
        <summary>Add-Description</summary>
            <p>
            ```
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
        ```
        </p>
    </details>

## Public
# Results
[Example results]
# Examples
[Example gif or movie]