function Add-WordChart {
    Param(
        [Parameter(Mandatory = $true)]
        [alias("CType")]
        [ValidateSet("Piechart", "Barchart")]
        [String] $chartType,
        [Parameter(Mandatory = $true)]
        [alias("CData")]
        $chartData,
        [Parameter(Mandatory = $true)]
        [alias("STitle")]
        [String] $sectionTitle,
        [Parameter(Mandatory = $true)]
        [alias("CTitle")]
        [String] $chartTitle

    )
    Add-WordText -WordDocument $reportFile -HeadingType Heading3 -Text $sectionTitle -Supress $true
    [array] $Names = foreach ($nameTMP in $chartData) {
        "$($nameTMP.Name) - [$($nameTMP.Values)]"
    }
    if ($chartType -like "*PieChart*") {    
        Add-WordPieChart -WordDocument $reportFile -ChartName $chartTitle -ChartLegendPosition Bottom -ChartLegendOverlay $false -Names $Names -Values $([array]$chartData.Values)
    }
    else {
        Add-WordBarChart -WordDocument $reportFile -ChartName $chartTitle -ChartLegendPosition Bottom -ChartLegendOverlay $false -Names $Names -Values $([array]$chartData.Values) -BarDirection Column   
    }
}