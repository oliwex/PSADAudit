function Get-GraphImage {
    Param(
        [Parameter(Mandatory = $true)]
        [Alias("GraphRoot")]
        $root, 
        [Alias("GraphLeaf")]
        $leaf, 
        [Alias("BasePathToGraphImage")]
        $pathToImage
    )

    $imagePath = Join-Path -Path $pathToImage -ChildPath "$root.png"
        
    $graphTMP = graph g {
        edge -from $root -To $leaf
    }
    
    $vizPath = Join-Path -Path $pathToImage -ChildPath "$root.vz"
    Set-Content -Path $vizPath -Value $graphTMP
    Export-PSGraph -Source $vizPath -Destination $imagePath

    #cleaning
    Remove-Item -Path $vizPath

    $imagePath
}