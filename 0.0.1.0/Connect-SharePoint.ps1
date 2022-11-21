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
# SIG # Begin signature block
# MIIFYQYJKoZIhvcNAQcCoIIFUjCCBU4CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQULZ/sjOnfL/wSdqKpM6brs7MQ
# uNqgggMAMIIC/DCCAeSgAwIBAgIQFnL4oVNG56NIRjNfzwNXejANBgkqhkiG9w0B
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
# DAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUIF4wkTOvsvwD7N/B7oKiB4uX
# EicwDQYJKoZIhvcNAQEBBQAEggEAzOJ9IlGLRHmjm/REU4e85m7zqZdM1SauSZrh
# RkAGUEnCl9G5mNdOnpt61pKWtWQDBEby41svi2RFjnvJrwxaFa+kwXAYRk8yJXRN
# 75Dlk7wgV1nt+PTyiksXWAikxPNDJwatm5Y60UBxTJyjlvYDiiDqUeryDjR/oXWI
# gKFgJqwzWkQPSduE8dpte+Bj9ZbqnYa4u2oEUNldwOVOvfgzrmkPuK/a6dAWUpXZ
# Xvoz6ZXByNZodCg1OfftEEla3a0q7IvJnPq9WrX4k0J/xnFIyqO7COXXZUlbPpgE
# 63kLqjP+ju5/nJc1P3h2pmHWjNgrCvJindpy2WDrZfmIdwb7SQ==
# SIG # End signature block
