#ExternalHelp $PSScriptRoot\New-SharePointViewQuery-help.xml
function New-SharePointViewQuery {

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

        [Parameter()]
        [Switch]
        $GroupBy,

        [Parameter(Position = 2)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $GroupByField,

        [Parameter()]
        [Switch]
        $Collapse,

        [Parameter()]
        [Switch]
        $OrderBy,

        [Parameter(Position = 3)]
        [ValidateNotNullOrEmpty()]
        [System.String[]]
        $OrderByField,

        [Parameter()]
        [Switch]
        $Override,

        [Parameter()]
        [Switch]
        $UseIndexForOrderBy,

        [Parameter()]
        [Switch]
        $Where,

        [Parameter(Position = 4)]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $WhereXml,

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

    if (!($Where) -and !($GroupBy) -and !($OrderBy)) {
        $selected = 'Where', 'GroupBy', 'OrderBy' | Out-GridView -Title 'View' -OutputMode Multiple
    }

    $innerQuery = New-Object -TypeName System.Text.StringBuilder
    [void]$innerQuery.Append('<Query>')
    if ($Where -or ('Where' -in $selected)) {
        if (!($WhereXml)) {
            [void]$innerQuery.Append('<Where>')
            if ($suppress) { return 1 }
            else {
                $recurseLevel = 0
                $recurseCount = 0
                $recurseStack = [System.Collections.Generic.List[System.String]]::New()
                do {
                    $type = $script:Default.QueryElements | Out-GridView -Title 'Type' -OutputMode Single
                    if ($type -in @('And', 'Or')) {
                        $recurseLevel++
                        $recurseStack.Add($type)
                        [void]$innerQuery.Append("<${type}>")
                    }
                    else {
                        if ($recurseLevel -ne 0) { $recurseCount++ }
                        if ($type -eq 'Membership') {
                            $membership = @('SPWeb.AllUsers', 'SPGroup', 'SPWeb.Groups', 'CurrentUserGroups', 'SPWeb.Users') | Out-GridView -Title $type -OutputMode Single
                            $field = $((Get-PnPField -List $List -Connection $script:Config.Connection | Where-Object { $_.CanBeDeleted -eq $true }).InternalName | Out-GridView -Title 'Field' -OutputMode Single)
                            [void]$innerQuery.Append("<${type} Type='${membership}'><FieldRef Name='${field}' />")
                        }
                        else {
                            [void]$innerQuery.Append("<${type}>")
                            Switch ($type) {
                                'DateRangesOverlap' {
                                    $eventDate = $((Get-PnPField -List $List -Connection $script:Config.Connection | Where-Object { $_.CanBeDeleted -eq $true }).InternalName | Out-GridView -Title 'EventDate' -OutputMode Single)
                                    $endDate = $((Get-PnPField -List $List -Connection $script:Config.Connection | Where-Object { $_.CanBeDeleted -eq $true }).InternalName | Out-GridView -Title 'EndDate' -OutputMode Single)
                                    $recurrenceId = $((Get-PnPField -List $List -Connection $script:Config.Connection | Where-Object { $_.CanBeDeleted -eq $true }).InternalName | Out-GridView -Title 'RecurrenceID' -OutputMode Single)
                                    $value = 'Now', 'Today', 'Month' | Out-GridView -Title 'Value' -OutputMode Single
                                    [void]$innerQuery.Append("<FieldRef Name='${eventDate}' />")
                                    [void]$innerQuery.Append("<FieldRef Name='${endDate}' />")
                                    [void]$innerQuery.Append("<FieldRef Name='${recurrenceId}' />")
                                    [void]$innerQuery.Append("<Value Type='DateTime'><${value} /></Value>")
                                }
                                'In' {
                                    $field = $((Get-PnPField -List $List -Connection $script:Config.Connection | Where-Object { $_.CanBeDeleted -eq $true }).InternalName | Out-GridView -Title 'Field' -OutputMode Single)
                                    $values = $(([xml](Get-PnPField -Identity $field -List $List -Connection $script:Config.Connection).SchemaXml).Field.CHOICES.CHOICE | Out-GridView -Title 'Values' -OutputMode Multiple)
                                    [void]$innerQuery.Append("<FieldRef Name='${field}' /><Values>")
                                    ForEach ($v in $values) {
                                        [void]$innerQuery.Append("<Value Type='Text'>${v}</Value>")
                                    }
                                    [void]$innerQuery.Append("</Values>")
                                }
                                Default {
                                    $field = $((Get-PnPField -List $List -Connection $script:Config.Connection | Where-Object { $_.CanBeDeleted -eq $true }).InternalName | Out-GridView -Title 'Field' -OutputMode Single)
                                    [void]$innerQuery.Append("<FieldRef Name='${field}' />")
                                    if ($type -ne 'IsNull') {
                                        $value = Read-Host -Prompt 'Value'
                                        [void]$innerQuery.Append("<Value Type='Text'>${value}</Value>")
                                    }
                                }
                            }
                        }
                        [void]$innerQuery.Append("</${type}>")
                    }
                    if ($recurseCount -eq 2) {
                        $recurseLevel--
                        $recurseCount--
                        [void]$innerQuery.Append("</$($recurseStack[$($recurseStack.Count - 1)])>")
                    }
                } until (($recurseLevel -eq 0) -and ($recurseCount -eq 0))
            }
            [void]$innerQuery.Append('</Where>')
        }
        else {
            [void]$innerQuery.Append($WhereXml)
        }
    }
    if ($GroupBy -or ('GroupBy' -in $selected)) {
        [void]$innerQuery.Append('<GroupBy')
        if ($Collapse) {
            [void]$innerQuery.Append(' Collapse="TRUE"')
        }
        [void]$innerQuery.Append('>')
        if (!($GroupByField)) {
            if (!($suppress)) {
                $GroupByField = $((Get-PnPField -List $List -Connection $script:Config.Connection | Where-Object { $_.CanBeDeleted -eq $true }).InternalName | Out-GridView -Title 'Field' -OutputMode Single)
            }
            else {
                return 1
            }
        }
        [void]$innerQuery.Append("<FieldRef Name='${GroupByField}' /></GroupBy>")
    }
    if ($OrderBy -or ('OrderBy' -in $selected)) {
        [void]$innerQuery.Append('<OrderBy')
        if ($Override) { [void]$innerQuery.Append(' Override="TRUE"') }
        if ($UseIndexForOrderBy) { [void]$innerQuery.Append(' UseIndexForOrderBy="TRUE"') }
        [void]$innerQuery.Append('>')
        if (!($OrderByField)) {
            if (!($suppress)) {
                do {
                    $field = $((Get-PnPField -List $List -Connection $script:Config.Connection | Where-Object { $_.CanBeDeleted -eq $true }).InternalName | Out-GridView -Title 'Field' -OutputMode Single)
                    [void]$innerQuery.Append("<FieldRef Name='${field}' />")
                    $yes = [System.Management.Automation.Host.ChoiceDescription]::New("&Yes", 'Yes')
                    $no = [System.Management.Automation.Host.ChoiceDescription]::New("&No", 'No')
                    $choice = $Host.UI.PromptForChoice('Select Another?', 'OrderByField', @($yes, $no), 0)
                } until ($choice -eq 1)
            }
            else {
                return 1
            }
        }
        else {
            ForEach ($f in $OrderByField) {
                [void]$innerQuery.Append("<FieldRef Name='${f}' />")
            }
        }
        [void]$innerQuery.Append('</OrderBy>')
    }
    [void]$innerQuery.Append('</Query>')
    $innerQuery.ToString()
}
# SIG # Begin signature block
# MIIFYQYJKoZIhvcNAQcCoIIFUjCCBU4CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU7/lugFPASfmvyZVJ4fYBK4j4
# q7mgggMAMIIC/DCCAeSgAwIBAgIQFnL4oVNG56NIRjNfzwNXejANBgkqhkiG9w0B
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
# DAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUO31W8gxnURrGuLm7ULI62Qs4
# 9XYwDQYJKoZIhvcNAQEBBQAEggEA6vTncYC/8Q1JRj5hi94fLqTF0JC9aDhGKyX3
# bKxuA9xwM0Gq32r+XLBu0ek07G0kxnw6vSLEusK9xhbMW4mIrahYQDrJyJNm/G8E
# 3rPExedZdKkkxgP7Vbj8kuoJLNO42jfYRP1G1Enypo/1xEDK2DDmLjZjWsCLNrXt
# f9ghHYxU98WimSEpqZ8eDM34mITMuE2h10FjtNQlW5iv5hHY4a3VqUeWH8rPJaO8
# 088vbtRXh3AppDoXoljWeFvuC+zKcq3omVy7zHbM34HKtCumee2gWQomWZ4NmA4g
# k3AayrTVgTRUC4icIoQgvO36M3zJCiNqUCNL24JxEAzsKUvf0Q==
# SIG # End signature block
