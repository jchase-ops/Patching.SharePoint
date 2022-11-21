#ExternalHelp $PSScriptRoot\Set-SharePointView-help.xml
function Set-SharePointView {

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
        $View,

        [Parameter(Mandatory, Position = 3)]
        [ValidateNotNullOrEmpty()]
        [System.Collections.Hashtable]
        $Values,

        [Parameter(Position = 4)]
        [ValidateNotNullOrEmpty()]
        [System.String[]]
        $Fields,

        [Parameter(Position = 5)]
        [ValidateNotNullOrEmpty()]
        [System.Collections.Hashtable]
        $Aggregations,

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

    if (!($View)) {
        if (!($List)) {
            if (!($suppress)) {
                $List = $((Get-PnPList -Connection $script:Config.Connection).Title | Out-GridView -Title "SharePoint Lists" -OutputMode Single)
            }
            else {
                return 1
            }
        }
        $View = $((Get-PnPView -List $List -Connection $script:Config.Connection).Title | Out-GridView -Title 'Views' -OutputMode Single)
    }

    $params = @{
        Identity = $View
        Values = $Values
        Connection = $script:Config.Connection
    }
    if ($List) { $params.List = $List }
    if ($Fields) { $params.Fields = $($Fields -Join ',') }
    if ($Aggregations) {
        $string = New-Object -TypeName System.Text.StringBuilder
        ForEach ($key in $Aggregations.Keys) {
            [void]$string.Append("<FieldRef Name='${key}' Type='$($Aggregations.$key)' />")
        }
        $params.Aggregations = $string.ToString()
    }
    if (!($suppress)) {
        Write-Host 'Setting view...' -NoNewline
    }
    Set-PnPView @params
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
# SIG # Begin signature block
# MIIFYQYJKoZIhvcNAQcCoIIFUjCCBU4CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUv+qBIjVt9ANoW7DmJmRu+Jra
# r9SgggMAMIIC/DCCAeSgAwIBAgIQFnL4oVNG56NIRjNfzwNXejANBgkqhkiG9w0B
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
# DAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUlfDppos6XoH/e8y2Le5BJDJy
# ku8wDQYJKoZIhvcNAQEBBQAEggEApIkjMt9FrqhcxRVsWNggtwcuMtxHb+1Wcrkm
# s2LWygDNwWTH+FAtst0oJrbyI8L2ElFcvMM+w2GE9vnVtmKcZru8fB/rvv4tq9tc
# 7g73ofBJRFg7s9gawkU69m10oFsRkPIYbKqu+wWcBUvmcuGDTFYDC0X8z+iIqe79
# Zm2ePF2cgyM7JopMtMRbAuFbG3Wk1veLXdEcTmNLKih9b8ySfgnM3kwUsFebjpt8
# cMxJf+e4EMqLfb8pBc0Nt/x43Db64u58EDR8278ss9D7ykCTZigx01rYX3IykqLy
# BXoZ/ATV8xdfeoviTEkB9TwbXPhTh78S4Kh+R1SQlD1SZlUWDw==
# SIG # End signature block
