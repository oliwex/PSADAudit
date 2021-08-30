function Get-GraphImage {
    Param(
        [Alias("GraphRoot")]
        $root,
        [Alias("GraphMiddle")]
        $middle, 
        [Alias("GraphLeaf")]
        $leaf,
        [Alias("BasePathToGraphImage")]
        $pathToImage
    )

    $imagePath = Join-Path -Path $pathToImage -ChildPath "$middle.png"
    $graphTMP=$null
    if ($null -eq $root) #not have boss
    {
        $graphTMP = graph g {
            edge -From $middle -To $leaf
        }    
    }
    elseif ($null -eq $leaf) #not have employees below
    {
        $graphTMP = graph g {
            edge -From $root -To $middle
        } 
    }
    elseif (($null -eq $leaf) -and ($null -eq $root)) #not have boss and employes
    {
        Add-WordText -WordDocument $reportFile -Text "No Boss no DirectReports" -Supress $true      
    }
    else #have boss and employees
    {
        $graphTMP = graph g {
                    edge -From $root -To $middle
                    edge -From $middle -To $leaf
                }
    }
    
    $vizPath = Join-Path -Path $pathToImage -ChildPath "$middle.vz"
    Set-Content -Path $vizPath -Value $graphTMP
    Export-PSGraph -Source $vizPath -Destination $imagePath

    #cleaning
    Remove-Item -Path $vizPath

    $imagePath
}


