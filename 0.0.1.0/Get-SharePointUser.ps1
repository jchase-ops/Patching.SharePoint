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
