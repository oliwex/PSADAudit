$basePath = "C:\reporty\"
$graphFolders = @{
    GPO       = "GPO_Graph\"
    OU        = "OU_Graph\"
    FGPP      = "FGPP_Graph\"
    GROUP     = "GROUP_Graph\"
    USERS     = "USERS_Graph\"
    COMPUTERS = "COMPUTERS_Graph\"
}

function Get-ReportFolders {
    Param(
        [Parameter(Mandatory = $true)]
        [Alias("BasePath")]
        [string]$reportPath,
        [Alias("GraphFoldersHashtable")]
        $graphFolders
    )
    $graphFoldersOutput=@{}
    $($graphFolders.Keys) | ForEach-Object {
        $folderPath = Join-Path -Path $reportPath -ChildPath $_
        $graphFoldersOutput.Add($_, $folderPath)
        New-Item -Path $folderPath -ItemType Directory
    }
    $graphFoldersOutput
}
Get-ReportFolders -BasePath $basePath -GraphFoldersHashtable $graphFolders