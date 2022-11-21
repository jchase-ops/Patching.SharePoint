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
# SIG # Begin signature block
# MIIFYQYJKoZIhvcNAQcCoIIFUjCCBU4CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUxPuBtgb9DyYIyeqT8f700a9L
# bgegggMAMIIC/DCCAeSgAwIBAgIQFnL4oVNG56NIRjNfzwNXejANBgkqhkiG9w0B
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
# DAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQU+xSe1BhRxTC5GJ2Br59iXusU
# G48wDQYJKoZIhvcNAQEBBQAEggEAhhDBRZattyx5jNjZINSVbqIzomdOj6XbNg0P
# TCt0AYa7gKyUm5JWgxZvwXjczUhhAbUvb4sWaaxG+E7bCqHEMPFRhcbSsYl1GENu
# eY7n90DkuqG/rFCeBvMy7jZAObA8WPxVNYk2Ht8QE2WJqVGfy5udbZoo/yCoZo9U
# MOhh9CHjwKUdIDMFn3Vp26DWMEFc5AI+RCaVx+Zii2QaLgQ4ZaR7X1y59uDLfCbu
# Ldo9EJAvdgmSJMsxilLLPXnjY7wPYwKkqifHQTwZjtdzkgxIfJ3drS6uZsPjvzU7
# WAMLa6eJq/vwoKcnpJD2SHkNMw4PHsr85E7FkIcdRJLFKIRatg==
# SIG # End signature block
