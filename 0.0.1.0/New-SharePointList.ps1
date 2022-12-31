# .ExternalHelp $PSScriptRoot\New-SharePointList-help.xml
function New-SharePointList {

    [CmdletBinding()]

    Param (

        [Parameter(Position = 0)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Site,

        [Parameter(Mandatory, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Title,

        [Parameter(Position = 2)]
        [ValidateScript({ $_ -in $script:Default.Templates })]
        [System.String]
        $Template = 'GenericList',

        [Parameter()]
        [Switch]
        $EnableContentTypes,

        [Parameter()]
        [Switch]
        $EnableVersioning,

        [Parameter()]
        [Switch]
        $Hidden,

        [Parameter()]
        [Switch]
        $OnQuickLaunch,

        [Parameter()]
        [Switch]
        $Quiet
    )

    $suppress = if (!(Test-WindowVisible) -or $Quiet) { $true } else { $false }

    if (!($Site)) {
        if ($null -eq $script:Config.Connection) {
            try {
                $script:Config.Connection = Get-PnPConnection
            }
            catch {
                if ($suppress) {
                    Connect-SharePoint -Quiet
                }
                else {
                    Connect-SharePoint
                }
            }
        }
        $Site = ($script:Config.Connection.Url.Split('/') | Select-Object -Last 1)
    }
    else {
        if ($null -eq $script:Config.Connection) {
            if (($script:Config.Url -Replace '^.*/') -ne $Site) {
                $script:Config.Url = $script:Config.Url.Replace($($script:Config.Url -Replace '^.*/'), $Site)
            }
            if ($suppress) {
                Connect-SharePoint -Quiet
            }
            else {
                Connect-SharePoint
            }
        }
        else {
            if (($script:Config.Connection.Url -Replace '^.*/') -ne $Site) {
                $script:Config.Url = $script:Config.Url.Replace($($script:Config.Url -Replace '^.*/'), $Site)
                if ($suppress) {
                    Connect-SharePoint -Quiet
                }
                else {
                    Connect-SharePoint
                }
            }
        }
    }

    $params = @{
        Connection = $script:Config.Connection
        Title = $PSBoundParameters.Title
        Template = $Template
    }
    if ($EnableContentTypes) { $params.EnableContentTypes = $true }
    if ($EnableVersioning) { $params.EnableVersioning = $true }
    if ($Hidden) { $params.Hidden = $true }
    if ($OnQuickLaunch) { $params.OnQuickLaunch = $true }

    if (!($suppress)) {
        Write-Host 'Creating list...' -NoNewline
    }
    New-PnPList @params
    if ($?) {
        if (!($suppress)) {
            Write-Host 'Complete' -ForegroundColor Green
        }
    }
    else {
        if (!($suppress)) {
            Write-Host 'Failed' -ForegroundColor Red
        }
        else {
            return 2
        }
    }
}
