# .ExternalHelp $PSScriptRoot\New-SharePointListItem-help.xml
function New-SharePointListItem {

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

        [Parameter(Position = 2, ParameterSetName = 'Standard')]
        [Parameter(Position = 2, ParameterSetName = 'Batch')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Folder,

        [Parameter(Mandatory, Position = 3)]
        [ValidateNotNullOrEmpty()]
        [System.Collections.Hashtable[]]
        $Values,

        [Parameter(Position = 4, ParameterSetName = 'Standard')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Label,

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
            if ($Folder) { $params.Folder = $Folder }
            if ($Label) { $params.Label = $Label }
            
            if (!($suppress)) {
                Write-Host 'Creating items...' -NoNewline
            }
            ForEach ($v in $Values) {
                $params.Values = $v
                Add-PnPListItem @params
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
            if ($Folder) { $params.Folder = $Folder }
            if (!($suppress)) {
                Write-Host 'Creating batch...' -NoNewline
            }
            ForEach ($v in $Values) {
                $params.Values = $v
                Add-PnPListItem @params
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
