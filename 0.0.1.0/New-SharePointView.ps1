#ExternalHelp $PSScriptRoot\New-SharePointView-help.xml
function New-SharePointView {

    [CmdletBinding(DefaultParameterSetName = 'Default')]

    Param (

        [Parameter(Position = 0, ParameterSetName = 'Default')]
        [Parameter(Position = 0, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Site,

        [Parameter(Position = 1, ParameterSetName = 'Default')]
        [Parameter(Position = 1, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $List,

        [Parameter(Mandatory, Position = 2)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Title,

        [Parameter(Position = 3, ParameterSetName = 'Default')]
        [Parameter(Position = 3, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Query,

        [Parameter(Position = 4, ParameterSetName = 'Default')]
        [Parameter(Position = 4, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String[]]
        $Fields,

        [Parameter(Position = 5, ParameterSetName = 'Default')]
        [Parameter(Position = 5, ParameterSetName = 'Xml')]
        [ValidateSet('None', 'Html', 'Grid', 'Recurrence', 'Chart', 'Calendar', 'Gantt')]
        [System.String]
        $ViewType,

        [Parameter(Position = 6, ParameterSetName = 'Default')]
        [Parameter(Position = 6, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.Int32]
        $RowLimit,

        [Parameter(Position = 7, ParameterSetName = 'Default')]
        [Parameter(Position = 7, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.Collections.Hashtable[]]
        $Aggregations,

        [Parameter(Position = 8, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.Int16]
        $BaseViewID,

        [Parameter(Position = 9, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $ContentTypeID,

        [Parameter(Position = 10, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $CssStyleSheet,

        [Parameter(Position = 11, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $DisplayName,

        [Parameter(Position = 12, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $ImageUrl,

        [Parameter(Position = 13, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.Int16]
        $MobileItemLimit,

        [Parameter(Position = 14, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $MobileUrl,

        [Parameter(Position = 15, ParameterSetName = 'Xml')]
        [ValidateSet('HideUnapproved', 'Contributor', 'Moderator')]
        [System.String]
        $ModerationType,

        [Parameter(Position = 16, ParameterSetName = 'Xml')]
        [ValidateScript({ $_ -in $script:Default.PageTypes })]
        [System.String]
        $PageType,

        [Parameter(Position = 17, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Path,

        [Parameter(Position = 18, ParameterSetName = 'Xml')]
        [ValidateSet('FilesOnly', 'Recursive', 'RecursiveAll')]
        [System.String]
        $Scope,

        [Parameter(Position = 19, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $SetupPath,

        [Parameter(Position = 20, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $TargetId,

        [Parameter(Position = 21, ParameterSetName = 'Xml')]
        [ValidateSet('List', 'ContentType')]
        [System.String]
        $TargetType,

        [Parameter(Position = 22, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $ToolbarTemplate,

        [Parameter(Position = 23, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Url,

        [Parameter(Position = 24, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.Int16]
        $WebPartOrder,

        [Parameter(Position = 25, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $WebPartZoneID,

        [Parameter(ParameterSetName = 'Default')]
        [Switch]
        $Personal,

        [Parameter(ParameterSetName = 'Default')]
        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $DefaultView,

        [Parameter(ParameterSetName = 'Default')]
        [Switch]
        $Paged,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $AggregateView,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $DefaultViewForContentType,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $FailIfEmpty,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $FileDialog,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $FPModified,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $FreeForm,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $Hidden,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $IncludeRootFolder,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $MobileDefaultView,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $MobileView,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $OrderedView,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $ReadOnly,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $RecurrenceRowset,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $ReqAuth,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $RequiresClientIntegration,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $ShowHeaderUI,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $TabularView,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $Threaded,

        [Parameter(ParameterSetName = 'Default')]
        [Parameter(ParameterSetName = 'Xml')]
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

    if (!($Query)) {
        $Query = New-SharePointViewQuery -List $List
    }

    if (!($Fields)) {
        $Fields = $((Get-PnPField -List $List -Connection $script:Config.Connection).Title | Out-GridView -Title 'Fields' -OutputMode Multiple)
    }

    Switch ($PSCmdlet.ParameterSetName) {
        'Default' {
            $params = @{
                List       = $List
                Title      = $Title
                Query      = $Query
                Fields     = $Fields
                Connection = $script:Config.Connection
            }
            if ($ViewType) { $params.ViewType = $ViewType }
            if ($RowLimit) { $params.RowLimit = $RowLimit }
            if ($Personal) { $params.Personal = $true }
            if ($SetAsDefault) { $params.SetAsDefault = $true }
            if ($Aggregations) {
                $string = New-Object -TypeName System.Text.StringBuilder
                ForEach ($key in $Aggregations.Keys) {
                    [void]$string.Append("<FieldRef Name='${key}' Type='$($Aggregations.$key)' />")
                }
                $params.Aggregations = $string.ToString()
            }
            if (!($suppress)) {
                Write-Host 'Creating view...' -NoNewline
            }
            Add-PnPView @params
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
        'Xml' {
            $viewXml = New-Object -TypeName System.Text.StringBuilder
            [void]$viewXml.Append("<View Name='${Title}'")
            ForEach ($key in $($PSBoundParameters.Keys | Where-Object { $_ -notin @('Site', 'List', 'Title', 'Query') })) {
                if ($($PSBoundParameters[$key]).GetType().Name -ne 'SwitchParameter') {
                    if ($($PSBoundParameters[$key]).GetType().BaseType.Name -eq 'Array') {
                        [void]$viewXml.Append(" ${key}='$($PSBoundParameters[$key] -Join ',')'")
                    }
                    else {
                        [void]$viewXml.Append(" ${key}='$($PSBoundParameters[$key])'")
                    }
                }
                else {
                    [void]$viewXml.Append(" ${key}='TRUE'")
                }
            }
            [void]$viewXml.Append(">${Query}</View>")
            $params = @{
                List = $List
                ViewXml = $viewXml.ToString()
                Connection = $script:Config.Connection
            }
            if (!($suppress)) {
                Write-Host 'Creating view...' -NoNewline
            }
            Add-PnPViewsFromXML @params
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
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU9zS35zGI5akw4XiIHIH4Nmww
# +I+gggMAMIIC/DCCAeSgAwIBAgIQFnL4oVNG56NIRjNfzwNXejANBgkqhkiG9w0B
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
# DAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQU/1Sz/VLI0zdBb1VtbW96aReG
# EU8wDQYJKoZIhvcNAQEBBQAEggEA3vZleLLjavJ54H5EnAM7T1N4I6Zmg+9twMus
# FjbBBVjRe/+hS3/0Ad5ZHNYxPxZ2N7evF7vnZoURBaoUrGW4vR+V4Lway2cvj7fS
# UgW2HvqIIW2N3wxGYkRxsj+saiNRLhY+ZQD6Jp8X0w9KLTSJ53E76dSxtZnX+tlC
# 7Bu3kx8YizTb50HlN879x//dw4dG1EzCsE8gHY/20K8OGtRE2+rb+U9VC9CzLo1o
# AioKmd7aBO1GOZh4jhoQUNTK4+Nw4AgPNoJ92qt9e6A/nBYTnMP9h6aMymiZyH4D
# bP8gvaMciu8qOf2idIyOk4dfq0VgEnb5KVwhmdil9gGUFMhBvg==
# SIG # End signature block
