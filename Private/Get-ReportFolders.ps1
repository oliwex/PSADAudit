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
        $folderPath = Join-Path -Path $reportPath -ChildPath $graphFolders[$_]
        New-Item -Path $folderPath -ItemType Directory | Out-Null
        $graphFoldersOutput.Add($_, $folderPath)
    }
    $graphFoldersOutput
}

