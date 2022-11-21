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
# SIG # Begin signature block
# MIIFYQYJKoZIhvcNAQcCoIIFUjCCBU4CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUeGgbruN2m+reM+ehfqBPGiWJ
# rR+gggMAMIIC/DCCAeSgAwIBAgIQFnL4oVNG56NIRjNfzwNXejANBgkqhkiG9w0B
# AQUFADAWMRQwEgYDVQQDDAtDZXJ0LTAzNDU2MDAeFw0yMTEyMDIwNDU5MTJaFw0y
# MjEyMDIwNTE5MTJaMBYxFDASBgNVBAMMC0NlcnQtMDM0NTYwMIIBIjANBgkqhkiG
# 9w0BAQEFAAOCAQ8AMIIBCgKCAQEA8daSAcUBI0Xx8sMMlSpsCV+24lY46RsxX8iC
# bB7ZM19b/GBjwMo0TCb28ssbZ/P8liNJICrSbyIkQDrIrjqtAdyAPdPAYHONTHad
# 0fuOQQT5MkO5HAxUYLz/6H/xq92lKQFxz5Wgzw+3KOyignY8V8ZZ379z/WqQbNCV
# +29zb9YWOK7eXQ9x8s4+SOizqUE3zkOuijf86I9vZmzMYhsxE7if0R0UlQsLlvTA
# kH/m4IjHem8rl/kC+O71lU7l9475XrUUR3Fxebqh9YoCEZh2eE81TLQcnvK8zgqP
# F+X4INdNPD6zO4T1Nbz0Ccev7mj37+pk/eL5R5aV+NJgqAzhvQIDAQABo0YwRDAO
# BgNVHQ8BAf8EBAMCBaAwEwYDVR0lBAwwCgYIKwYBBQUHAwMwHQYDVR0OBBYEFFNN
# e4x6JSqbcnTR354fVSEgQ0VYMA0GCSqGSIb3DQEBBQUAA4IBAQBXfA8VgaMD2c/v
# Sv8gnS/LWri51BBqcUFE9JYMxEIzlEt2ZfJsG+INaQqzBoyCDx/oMQH7wdFRvDjQ
# QsXpNTo7wH7WytFe9KJrOz2uGG0EnIYHK0dTFIMVOcM9VsWWPG40EAzD//55xX/d
# pBL+L4SSTujbR3ptni8Agu5GiRhTpxwl1L/HLC2QYYMoUKiAxL1p61+cHRj6wMzl
# jxnrMIcBhKioaXnwWdKPCN66Jk8IYdqr8afcRYiwtDi+8Hk2/9nB9HwPox3Dtf8H
# jH0O2/8NiJTeOBFSfrWPM9r4j4NWR8IuLwsqHUfXJEQa9SOxhHvxaNMR/Fhq1GVj
# qUClZiXiMYIByzCCAccCAQEwKjAWMRQwEgYDVQQDDAtDZXJ0LTAzNDU2MAIQFnL4
# oVNG56NIRjNfzwNXejAJBgUrDgMCGgUAoHgwGAYKKwYBBAGCNwIBDDEKMAigAoAA
# oQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4w
# DAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUwvIGSluh0RmG+V5zC6pb8ERl
# g2IwDQYJKoZIhvcNAQEBBQAEggEA2VECaQfHGLO7KKxHPM4HYgW4ufkHH8XvWzG9
# EFpJZYFAU4nZ/RolzS67fnmQGdXw3fdDm/CsAI8TkiGP5RE/clNJwCJlsVWuKJ00
# sCLsmzHJhb0aHF+JZEEqZCle7j4/Cddak7vHkE+2r1Dc+TA7jnT6O3YQuT5x6CfV
# J0huEp4w0I57N6eAuVdCJi+pwlyZ2FVmIPkgMVZmTmtM33/eIDsoyxGPHtJgv32j
# 5+DK59JMD+mTPmrGRpNWOpI+9DpIIJ4hvtZia/5FAPofsF0HV4ttoXu+PRKshmUT
# kX022R0Iz6MFG5tXNlFFyinetxhWNUhp3EGl726+BxGEptIDLQ==
# SIG # End signature block
