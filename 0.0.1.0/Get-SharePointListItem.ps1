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
# SIG # Begin signature block
# MIIFYQYJKoZIhvcNAQcCoIIFUjCCBU4CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUhz4+SsU+zCXMMoo5+1m0XXsS
# GPugggMAMIIC/DCCAeSgAwIBAgIQFnL4oVNG56NIRjNfzwNXejANBgkqhkiG9w0B
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
# DAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUMvpyoNpHwRaZmlDcRMFmh4zf
# 2ecwDQYJKoZIhvcNAQEBBQAEggEAIoNIDTdY1HSOx7wo+NBUBwRIPkKSJkK7NoS+
# mMP+TmQCfKK8wm3dFqPIPFpfDzC+rHn+ftJ2qaDvD5qRlh0Ye9loYw5ewBbaf5tk
# P1wE8/x1onNWzIxt0CFPiXyGlNXhKqz/04CldFZ6UEYNB9xHpd47KHekjni59feb
# Meck8d9JSkUp+hkERaCfOk1r6D5ar3fXxagj3ishtttp3TlGh0sVBXjg8cJIrC1G
# n1kkki1HFQgr1oUTv3L2a2fhZjnVvr22pgY/xwh/mFY4+SY/lB8RvIGvutLKx4Rf
# 3pq+4SCWyVJSzPBAY2dw3JM1k/ndx3cdvmf/O/jY/tHg/jnibg==
# SIG # End signature block
