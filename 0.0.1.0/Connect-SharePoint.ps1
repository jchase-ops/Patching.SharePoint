# .ExternalHelp $PSScriptRoot\Connect-SharePoint-help.xml
function Connect-SharePoint {

    [CmdletBinding()]

    Param (

        [Parameter(Position = 0)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $TenantName,

        [Parameter(Position = 1)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Site,

        [Parameter()]
        [Switch]
        $PassThru,

        [Parameter()]
        [Switch]
        $Quiet
    )

    $suppress = if (!(Test-WindowVisible) -or $Quiet) { $true } else { $false }

    if (!($TenantName)) {
        if ($null -eq $script:Config.Url) {
            if (!($suppress)) {
                $TenantName = Read-Host -Prompt 'TenantName'
                $script:Config.Url = "https://${TenantName}.sharepoint.com/sites"
            }
            else {
                return 1
            }
        }
    }
    else {
        $script:Config.Url = "https://${TenantName}.sharepoint.com/sites"
    }

    if (!($Site)) {
        if ($script:Config.Url -like "*/sites") {
            if (!($suppress)) {
                $Site = Read-Host -Prompt 'Site'
                $script:Config.Url = "$($script:Config.Url)/${Site}"
            }
            else {
                return 1
            }
        }
    }
    else {
        if (($script:Config.Url -Replace '^.*/') -eq 'sites') {
            $script:Config.Url = "$($script:Config.Url)/${Site}"
        }
        else {
            if (($script:Config.Url -Replace '^.*/') -ne $Site) {
                $script:Config.Url = $script:Config.Url.Replace($($script:Config.Url -Replace '^.*/'), $Site)
            }
        }
    }

    $script:Config | Export-Clixml -Path "$PSScriptRoot\config.xml" -Depth 100

    $script:Config.Connection = Connect-PnPOnline -Url $script:Config.Url -Interactive -ReturnConnection
    if ($?) {
        if (!($suppress)) {
            Write-Host "Connected to " -NoNewline
            Write-Host $script:Config.Url -ForegroundColor Green
        }
    }
    else {
        $script:Config.Connection = $null
        if (!($suppress)) {
            Write-Host "Failed to connect to " -NoNewline
            Write-Host $script:Config.Url -ForegroundColor Yellow
        }
        else {
            return 2
        }
    }

    if ($PassThru -and ($null -ne $script:Config.Connection)) {
        $script:Config
    }
}
