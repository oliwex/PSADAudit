#Path to folders
$graphFolders = @{
    GPO       = "GPO_Graph\"
    OU        = "OU_Graph\"
    FGPP      = "FGPP_Graph\"
    GROUP     = "GROUP_Graph\"
    USERS     = "USERS_Graph\"
    COMPUTERS = "COMPUTERS_Graph\"
}

#Description json
$pathToDescription="$(($env:PSModulePath -split ";")[1])\PSADAudit\Private\Text\Description.json"
