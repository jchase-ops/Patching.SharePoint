#ExternalHelp $PSScriptRoot\Remove-SharePointItem-help.xml
function Remove-SharePointItem {

    [CmdletBinding(DefaultParameterSetName = 'List')]

    Param (

        [Parameter(Position = 0, ParameterSetName = 'Field')]
        [Parameter(Position = 0, ParameterSetName = 'List')]
        [Parameter(Position = 0, ParameterSetName = 'ListItem')]
        [Parameter(Position = 0, ParameterSetName = 'View')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Site,

        [Parameter(Position = 1, ParameterSetName = 'Field')]
        [Parameter(Position = 1, ParameterSetName = 'List')]
        [Parameter(Position = 1, ParameterSetName = 'ListItem')]
        [Parameter(Position = 1, ParameterSetName = 'View')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $List,

        [Parameter(Mandatory, Position = 2, ParameterSetName = 'Field')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Field,

        [Parameter(Mandatory, Position = 3, ParameterSetName = 'ListItem', ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $ListItem,

        [Parameter(Mandatory, Position = 4, ParameterSetName = 'View')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $View,

        [Parameter(ParameterSetName = 'List')]
        [Parameter(ParameterSetName = 'ListItem')]
        [Switch]
        $Recycle,

        [Parameter(ParameterSetName = 'Field')]
        [Parameter(ParameterSetName = 'List')]
        [Parameter(ParameterSetName = 'ListItem')]
        [Parameter(ParameterSetName = 'View')]
        [Switch]
        $Force,

        [Parameter(ParameterSetName = 'ListItem')]
        [Switch]
        $Batch,

        [Parameter(ParameterSetName = 'Field')]
        [Parameter(ParameterSetName = 'List')]
        [Parameter(ParameterSetName = 'View')]
        [Switch]
        $WhatIf,

        [Parameter(ParameterSetName = 'Field')]
        [Parameter(ParameterSetName = 'List')]
        [Parameter(ParameterSetName = 'ListItem')]
        [Parameter(ParameterSetName = 'View')]
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

    if (!($List)) {
        if (!($suppress)) {
            $List = $((Get-PnPList -Connection $script:Config.Connection).Title | Out-GridView -Title "SharePoint $Site Lists" -OutputMode Single)
        }
        else {
            return 1
        }
    }

    Switch ($PSCmdlet.ParameterSetName) {
        'Field' {
            $params = @{
                Identity = $Field
                List = $List
                Connection = $script:Config.Connection
            }
            if ($Force) { $params.Force = $true }
            if ($WhatIf) { $params.WhatIf = $true }
            if (!($suppress)) {
                Write-Host 'Removing field...' -NoNewline
            }
            Remove-PnPField @params
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
        'List' {
            $params = @{
                Identity = $List
                Connection = $script:Config.Connection
            }
            if ($Recycle) { $params.Recycle = $true }
            if ($Force) { $params.Force = $true }
            if ($WhatIf) { $params.WhatIf = $true }
            if (!($suppress)) {
                Write-Host 'Removing list...' -NoNewline
            }
            Remove-PnPList @params
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
        'ListItem' {
            $params = @{
                List = $List
                Connection = $script:Config.Connection
            }
            if ($Recycle) { $params.Recycle = $true }
            if ($Force) { $params.Force = $true }
            if ($Batch -or @($ListItem).Count -gt 25) {
                $pnpBatch = New-PnPBatch
                $params.Batch = $pnpBatch
            }
            if (!($suppress)) {
                Write-Host 'Removing ListItem...' -NoNewline
            }
            $ListItem | ForEach-Object {
                $params.Identity = $_
                Remove-PnPListItem @params
            }
            if ($Batch -or $($null -ne $pnpBatch)) {
                Invoke-PnPBatch -Batch $pnpBatch
            }
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
        'View' {
            $params = @{
                List = $List
                Identity = $View
                Connection = $script:Config.Connection
            }
            if ($Force) { $params.Force = $true }
            if ($WhatIf) { $params.WhatIf = $true }
            if (!($suppress)) {
                Write-Host 'Removing view...' -NoNewline
            }
            Remove-PnPView @params
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
    }
}
