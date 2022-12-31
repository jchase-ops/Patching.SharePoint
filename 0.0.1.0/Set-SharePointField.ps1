# .ExternalHelp $PSScriptRoot\Set-SharePointField-help.xml
function Set-SharePointField {

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
        $Field,

        [Parameter(Mandatory, Position = 3)]
        [ValidateNotNullOrEmpty()]
        [System.Collections.Hashtable]
        $Values,

        [Parameter()]
        [Switch]
        $UpdateExistingLists,

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
            $List = $((Get-PnPList -Connection $script:Config.Connection).Title | Out-GridView -Title "SharePoint Lists" -OutputMode Single)
        }
        else {
            return 1
        }
    }

    if (!($Field)) {
        if (!($suppress)) {
            $Field = $((Get-PnPField -List $List -Connection $script:Config.Connection | Where-Object { $_.CanBeDeleted -eq $true -and $_.TypeDisplayName -in @('Choice', 'MultiChoice') }).Title | Out-GridView -Title "$List Fields" -OutputMode Single)
        }
        else {
            return 1
        }
    }

    $params = @{
        List = $List
        Identity = $Field
        Connection = $script:Config.Connection
        Values = $Values
    }
    if ($UpdateExistingLists) { $params.$UpdateExistingLists = $true }
    
    if (!($suppress)) {
        Write-Host 'Setting field...' -NoNewline
    }
    Set-PnPField @params
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
