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
# SIG # Begin signature block
# MIIFYQYJKoZIhvcNAQcCoIIFUjCCBU4CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU+iYPFgUQBOZhTDSABTzdUN4/
# 0SGgggMAMIIC/DCCAeSgAwIBAgIQFnL4oVNG56NIRjNfzwNXejANBgkqhkiG9w0B
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
# DAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUuBKHrF0noE7zwbEEZfyZIXbP
# K4MwDQYJKoZIhvcNAQEBBQAEggEAKjjWuGaOIMtpJgTXSzG9ebDiZuNIN7k79b0E
# xWTxzuY2DL3drij2nkn3hvkfK1UKdT79bljlEimLgLvZ4V/+SmacbLZu8isbseeV
# PsG2QdM5TwBP0+SRrKDI+V0dse+T1GbmUGibbQFnP2PWv4JA/qbLQJjZiWnmLoG5
# 7+YLf8fjWHkPc136T7mV3c9uA9MoDchoZ+nJ3MLpmppUIDPYtmUoHa7FZtBKOWjM
# w1O9QFp23HshE5Ag4+FKcufj5xl+Xn/0TR/rp7bzTivjuXKHFZIWOjWtsZlr6mn3
# 42NOAPKaRdRvZyTrdc1p8PdrhYwdCMo4OdTiTbEfATn+kX8FvA==
# SIG # End signature block
