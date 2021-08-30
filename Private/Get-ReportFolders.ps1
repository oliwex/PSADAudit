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
    $graphFolders.Keys.clone() | ForEach-Object {
        $graphFolders[$_] = Join-Path -Path $reportPath -ChildPath $_
    }
    $graphFolders
}
