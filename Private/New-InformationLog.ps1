function New-InformationLog
{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true,HelpMessage="LogPath",Position=0)]
        [String]$logPath,
        [Parameter(Mandatory=$true,HelpMessage="Message",Position=1)]
        [String]$message,
        [Parameter(Mandatory=$true,HelpMessage="Color",Position=2)]
        [String]$color
        )
    $datetime=Get-Date -Format "HH.mm.ss.ffff_dd.MM.yyyy"
    "[$datetime] $message" >> $logPath
    Write-Host "[$datetime] $message " -ForegroundColor $color
}