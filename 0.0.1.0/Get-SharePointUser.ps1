# .ExternalHelp $PSScriptRoot\Get-SharePointUser-help.xml
function Get-SharePointUser {

    [CmdletBinding(DefaultParameterSetName = 'SamAccountName')]

    Param (

        [Parameter(Position = 0, ParameterSetName = 'SamAccountName')]
        [Parameter(Position = 0, ParameterSetName = 'UserPrincipalName')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Site,

        [Parameter(Position = 1, ParameterSetName = 'SamAccountName', ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $SamAccountName,

        [Parameter(Position = 2, ParameterSetName = 'SamAccountName')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Domain,

        [Parameter(Position = 3, ParameterSetName = 'SamAccountName')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $EmailSuffix,

        [Parameter(Mandatory, Position = 1, ValueFromPipeline, ParameterSetName = 'UserPrincipalName')]
        [ValidatePattern("^[a-zA-Z]{1,}\.[a-zA-Z]{1,}@[a-zA-Z]{1,}\.(com|edu|gov)$")]
        [System.String]
        $UserPrincipalName,

        [Parameter(ParameterSetName = 'SamAccountName')]
        [Parameter(ParameterSetName = 'UserPrincipalName')]
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
        $Site = ($script:Config.Connection.Url.Split('/') | Select-Object -Last 1)
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

    Switch ($PSCmdlet.ParameterSetName) {
        'SamAccountName' {
            if (!($SamAccountName)) {
                $SamAccountName = Read-Host -Prompt 'UserName'
            }
            if (!($Domain)) {
                $Domain = Read-Host -Prompt 'Domain'
            }
            if (!($EmailSuffix)) {
                $EmailSuffix = Read-Host -Prompt 'Email Suffix'
            }
            $identity = "i:0#.f|membership|${SamAccountName}@${Domain}.${EmailSuffix}"
        }
        'UserPrincipalName' {
            $identity = "i:0#.f|membership|${UserPrincipalName}"
        }
    }
    Get-PnPUser -Identity $identity -Connection $script:Config.Connection
}
# SIG # Begin signature block
# MIIFYQYJKoZIhvcNAQcCoIIFUjCCBU4CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUR2T54agnAbsA4sg5CEjR+B2h
# FrigggMAMIIC/DCCAeSgAwIBAgIQFnL4oVNG56NIRjNfzwNXejANBgkqhkiG9w0B
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
# DAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUUkSrhQsd75d4yT1/bQEIWS8G
# UOMwDQYJKoZIhvcNAQEBBQAEggEAm4C/nIjnhsUjDEutznXidfTve6zdkJu1L1B4
# M1oFidM/hE7gDqgxONAwMZEWNunU4gt2ZpLMFsMMqNP1mbrer0/1To55udmN2m+Q
# DjqAuu4rxb4szF1Ka5IAyvsN8I/Fv2ABj+SqEJ/x53eOxXk8J6L2/GEkmrz95slj
# +biQ99bT9QlxxAx4Fto+atqXPZrdr1irNq7xlbuFRz9phK+P2dbmjWEQQEpHfRmm
# qpG0oxubBRFmD+73PAaI+bb7m/NkDQOupMffdb8L3QVqzMYkEa439YB2GBoRGEi2
# B8O1qJE5fQrP34eZjPXDw3TUxSvpSmMWi+wAptpPLngasOg66Q==
# SIG # End signature block
