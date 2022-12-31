#ExternalHelp $PSScriptRoot\Set-SharePointListItem-help.xml
function Set-SharePointListItem {

    [CmdletBinding(DefaultParameterSetName = 'Standard')]

    Param (

        [Parameter(Position = 0, ParameterSetName = 'Standard')]
        [Parameter(Position = 0, ParameterSetName = 'Batch')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Site,

        [Parameter(Position = 1, ParameterSetName = 'Standard')]
        [Parameter(Position = 1, ParameterSetName = 'Batch')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $List,

        [Parameter(Position = 2, ParameterSetName = 'Standard', ValueFromPipelineByPropertyName)]
        [Parameter(Position = 2, ParameterSetName = 'Batch', ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [System.String[]]
        $ID,

        [Parameter(Mandatory, Position = 3)]
        [ValidateNotNullOrEmpty()]
        [System.Collections.Hashtable[]]
        $Values,

        [Parameter(Position = 4, ParameterSetName = 'Standard')]
        [Parameter(Position = 4, ParameterSetName = 'Batch')]
        [ValidateSet('Update', 'SystemUpdate', 'UpdateOverwriteVersion')]
        [System.String]
        $UpdateType,

        [Parameter(Position = 5, ParameterSetName = 'Standard')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Label,

        [Parameter(ParameterSetName = 'Standard')]
        [Switch]
        $ClearLabel,

        [Parameter(ParameterSetName = 'Standard')]
        [Parameter(ParameterSetName = 'Batch')]
        [Switch]
        $Force,

        [Parameter(Mandatory, ParameterSetName = 'Batch')]
        [Switch]
        $Batch,

        [Parameter(ParameterSetName = 'Standard')]
        [Parameter(ParameterSetName = 'Batch')]
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
            $List = $((Get-PnPList -Connection $script:Config.Connection).Title | Out-GridView -Title "SharePoint Lists" -OutputMode Single)
        }
        else {
            return 1
        }
    }

    Switch ($PSCmdlet.ParameterSetName) {
        'Standard' {
            $params = @{
                List = $List
                Connection = $script:Config.Connection
            }
            if ($UpdateType) { $params.UpdateType = $UpdateType }
            if ($Label) { $params.Label = $Label }
            if ($ClearLabel) { $params.ClearLabel = $true }
            if ($Force) { $params.Force = $true }

            if (!($suppress)) {
                Write-Host 'Setting items...' -NoNewline
            }

            ForEach ($i in $ID) {
                $params.Identity = $i
                $params.Values = $Values[$($ID.IndexOf($i))]
                Set-PnpListItem @params
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
        'Batch' {
            $batch = New-PnPBatch -Connection $script:Config.Connection
            $params = @{
                List = $List
                Connection = $script:Config.Connection
                Batch = $batch
            }
            if ($UpdateType) { $params.UpdateType = $UpdateType }
            if ($Force) { $params.Force = $true }

            if (!($suppress)) {
                Write-Host 'Creating batch...' -NoNewline
            }
            ForEach ($i in $ID) {
                $params.Identity = $i
                $params.Values = $Values[$($ID.IndexOf($i))]
                Set-PnPListItem @params
            }
            if (!($suppress)) {
                Write-Host 'Complete'
                Write-Host 'Invoking batch...' -NoNewline
            }
            Invoke-PnPBatch $batch
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
