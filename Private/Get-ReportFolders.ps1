function Get-ReportFolders {
    Param(
        [Parameter(Mandatory = $true)]
        [Alias("BasePath")]
        [string]$reportPath,
        [Alias("GraphFoldersHashtable")]
        $graphFolders
    )

    foreach ($key in $($graphFolders.Keys)) {
        $graphPath = Join-Path -Path $reportPath -ChildPath $graphFolders[$key]
        $graphFolders[$key] = $graphPath
        New-Item -Path $graphPath -ItemType Directory
    }
    $graphFolders
}