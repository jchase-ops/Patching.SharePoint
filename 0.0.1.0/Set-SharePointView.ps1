#ExternalHelp $PSScriptRoot\Set-SharePointView-help.xml
function Set-SharePointView {

    [CmdletBinding()]

    Param (

        [Parameter(Position = 0)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Site,

        [Parameter(Position = 1)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $List,

        [Parameter(Position = 2)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $View,

        [Parameter(Mandatory, Position = 3)]
        [ValidateNotNullOrEmpty()]
        [System.Collections.Hashtable]
        $Values,

        [Parameter(Position = 4)]
        [ValidateNotNullOrEmpty()]
        [System.String[]]
        $Fields,

        [Parameter(Position = 5)]
        [ValidateNotNullOrEmpty()]
        [System.Collections.Hashtable]
        $Aggregations,

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
        $Site = ($script:Config.Connection.Url.Split('/') | Select-Object -Last 1).ToUpper()
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

    if (!($View)) {
        if (!($List)) {
            if (!($suppress)) {
                $List = $((Get-PnPList -Connection $script:Config.Connection).Title | Out-GridView -Title "SharePoint Lists" -OutputMode Single)
            }
            else {
                return 1
            }
        }
        $View = $((Get-PnPView -List $List -Connection $script:Config.Connection).Title | Out-GridView -Title 'Views' -OutputMode Single)
    }

    $params = @{
        Identity = $View
        Values = $Values
        Connection = $script:Config.Connection
    }
    if ($List) { $params.List = $List }
    if ($Fields) { $params.Fields = $($Fields -Join ',') }
    if ($Aggregations) {
        $string = New-Object -TypeName System.Text.StringBuilder
        ForEach ($key in $Aggregations.Keys) {
            [void]$string.Append("<FieldRef Name='${key}' Type='$($Aggregations.$key)' />")
        }
        $params.Aggregations = $string.ToString()
    }
    if (!($suppress)) {
        Write-Host 'Setting view...' -NoNewline
    }
    Set-PnPView @params
    if ($?) {
        if (!($suppress)) {
            Write-Host 'Success' -ForegroundColor Green
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
