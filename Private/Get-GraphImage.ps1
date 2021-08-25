function Get-GraphImage {
    Param(
        [Parameter(Mandatory = $true)]
        [Alias("GraphRoot")]
        $root,
        [Alias("GraphMiddle")]
        $middle, 
        [Alias("GraphLeaf")]
        $leaf, 
        [Alias("BasePathToGraphImage")]
        $pathToImage
    )

    $imagePath = Join-Path -Path $pathToImage -ChildPath "$root.png"
    $graphTMP=$null
    if ($root -eq $null)
    {
        $graphTMP = graph g {
            edge -From $middle -To $leaf
        }    
    }
    else
    {
        $graphTMP = graph g {
            edge -From $root -To $middle
            edge -From $middle -To $leaf
        }        
    }
    
    $vizPath = Join-Path -Path $pathToImage -ChildPath "$root.vz"
    Set-Content -Path $vizPath -Value $graphTMP
    Export-PSGraph -Source $vizPath -Destination $imagePath

    #cleaning
    Remove-Item -Path $vizPath

    $imagePath
}

 