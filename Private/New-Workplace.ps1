function New-Workplace
{
    do
    {
        $workplacePath=Read-Host "Get Path for Report:"
        $isPathCorrect=(([string]::IsNullOrWhiteSpace($workplacePath)) -or (Test-Path -Path $workplacePath -PathType Container))
    }
    while ($isPathCorrect)

    New-Item -Path $workplacePath -ItemType Directory | Out-Null
    $workplacePath
}

