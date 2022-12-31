# .ExternalHelp $PSScriptRoot\Get-SharePointListItem-help.xml
function Get-SharePointListItem {

    [CmdletBinding(DefaultParameterSetName = 'All')]

    Param (

        [Parameter(Position = 0, ParameterSetName = 'All')]
        [Parameter(Position = 0, ParameterSetName = 'Id')]
        [Parameter(Position = 0, ParameterSetName = 'Query')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Site,

        [Parameter(Position = 1, ParameterSetName = 'All')]
        [Parameter(Position = 1, ParameterSetName = 'Id')]
        [Parameter(Position = 1, ParameterSetName = 'Query')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $List,

        [Parameter(Position = 2, ParameterSetName = 'All')]
        [Parameter(Position = 2, ParameterSetName = 'Id')]
        [ValidateNotNullOrEmpty()]
        [System.String[]]
        $Fields,

        [Parameter(Position = 3, ParameterSetName = 'All')]
        [Parameter(Position = 3, ParameterSetName = 'Query')]
        [ValidateNotNullOrEmpty()]
        [System.Int32]
        $PageSize,

        [Parameter(Position = 4, ParameterSetName = 'All')]
        [Parameter(Position = 4, ParameterSetName = 'Query')]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.ScriptBlock]
        $ScriptBlock,

        [Parameter(Mandatory, Position = 5, ParameterSetName = 'Query')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Query,

        [Parameter(Mandatory, Position = 6, ParameterSetName = 'Id')]
        [ValidateNotNullOrEmpty()]
        [System.Int32[]]
        $ID,

        [Parameter(Position = 7, ParameterSetName = 'All')]
        [Parameter(Position = 7, ParameterSetName = 'Query')]
        [System.String]
        $FolderServerRelativeUrl,

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

    if (!($List)) {
        if (!($suppress)) {
            $List = $((Get-PnPList -Connection $script:Config.Connection).Title | Out-GridView -Title "SharePoint $Site Lists" -OutputMode Single)
        }
        else {
            return 1
        }
    }

    Switch ($PSCmdlet.ParameterSetName) {
        'All' {
            $params = @{
                List = $List
                Connection = $script:Config.Connection
            }
            if ($FolderServerRelativeUrl) { $params.FolderServerRelativeUrl = $FolderServerRelativeUrl }
            if ($Fields) { $params.Fields = $Fields }
            if ($PageSize) { $params.PageSize = $PageSize }
            if ($ScriptBlock) { $params.ScriptBlock = $ScriptBlock }
            Get-PnPListItem @params
        }
        'Id' {
            $params = @{
                List = $List
                Connection = $script:Config.Connection
            }
            if ($Fields) { $params.Fields = $Fields }
            ForEach ($n in $ID) {
                $params.Id = $n
                Get-PnPListItem @params
            }
        }
        'Query' {
            $params = @{
                List = $List
                Connection = $script:Config.Connection
                Query = $Query
            }
            if ($PageSize) { $params.PageSize = $PageSize }
            if ($FolderServerRelativeUrl) { $params.FolderServerRelativeUrl = $FolderServerRelativeUrl }
            if ($ScriptBlock) { $params.ScriptBlock = $ScriptBlock }
            Get-PnPListItem @params
        }
    }
}
