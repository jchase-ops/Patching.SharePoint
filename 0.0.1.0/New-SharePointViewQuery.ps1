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
