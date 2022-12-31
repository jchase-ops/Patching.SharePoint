#ExternalHelp $PSScriptRoot\Set-SharePointList-help.xml
function Set-SharePointList {

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
        $Title,

        [Parameter(Position = 3)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Description,

        [Parameter(Position = 4)]
        [ValidateSet('Auto', 'NewExperience', 'ClassicExperience')]
        [System.String]
        $ListExperience,

        [Parameter(Position = 5)]
        [ValidateNotNullOrEmpty()]
        [System.Int32]
        $MajorVersions,

        [Parameter(Position = 6)]
        [ValidateNotNullOrEmpty()]
        [System.Int32]
        $MinorVersions,

        [Parameter(Position = 7)]
        [ValidateSet('Approver', 'Author', 'Reader')]
        [System.String]
        $DraftVersionVisibility,

        [Parameter(Position = 8)]
        [ValidateSet('AllUsersReadAccess', 'AllUsersReadAccessOnItemsTheyCreate')]
        [System.String]
        $ReadSecurity,

        [Parameter(Position = 9)]
        [ValidateSet('WriteAllItems', 'WriteOnlyMyItems', 'WriteNoItems')]
        [System.String]
        $WriteSecurity,

        [Parameter(Position = 10)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Path,

        [Parameter()]
        [Switch]
        $EnableContentTypes,

        [Parameter()]
        [Switch]
        $BreakRoleInheritance,

        [Parameter()]
        [Switch]
        $ResetRoleInheritance,

        [Parameter()]
        [Switch]
        $CopyRoleAssignments,

        [Parameter()]
        [Switch]
        $ClearSubscopes,

        [Parameter()]
        [Switch]
        $Hidden,

        [Parameter()]
        [Switch]
        $ForceCheckout,

        [Parameter()]
        [Switch]
        $EnableAttachments,

        [Parameter()]
        [Switch]
        $EnableFolderCreation,

        [Parameter()]
        [Switch]
        $EnableVersioning,

        [Parameter()]
        [Switch]
        $EnableMinorVersions,

        [Parameter()]
        [Switch]
        $EnableModeration,

        [Parameter()]
        [Switch]
        $NoCrawl,

        [Parameter()]
        [Switch]
        $ExemptFromBlockDownloadOfNonViewableFiles,

        [Parameter()]
        [Switch]
        $DisableGridEditing,

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

    $params = @{
        Identity = $List
        Connection = $script:Config.Connection
    }
    ForEach ($key in $($PSBoundParameters.Keys | Where-Object { $_ -notin @('Site', 'List', 'Quiet') })) {
        if ($($PSBoundParameters[$key]).GetType().Name -eq 'SwitchParameter') {
            $params[$key] = $true
        }
        else {
            $params[$key] = $PSBoundParameters[$key]
        }
    }

    if (!($suppress)) {
        Write-Host 'Setting list...' -NoNewline
    }
    Set-PnPList @params
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
