# .ExternalHelp $PSScriptRoot\New-SharePointField-help.xml
function New-SharePointField {

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

        [Parameter(Position = 2, ParameterSetName = 'Default')]
        [Parameter(Position = 2, ParameterSetName = 'Xml')]
        [ValidateScript({ $_ -in $script:Default.FieldTypes })]
        [System.String]
        $Type,

        [Parameter(Mandatory, Position = 3, ParameterSetName = 'Default')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $InternalName,

        [Parameter(Mandatory, Position = 3, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Name,

        [Parameter(Position = 4, ParameterSetName = 'Default')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $DisplayName,


        [Parameter(Position = 4, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Title,

        [Parameter(Position = 5, ParameterSetName = 'Default')]
        [Parameter(Position = 5, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Id = [System.Guid]::NewGuid().Guid,

        [Parameter(Position = 6, ParameterSetName = 'Default')]
        [Parameter(Position = 6, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Formula,

        [Parameter(Position = 7, ParameterSetName = 'Default')]
        [Parameter(Position = 7, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String[]]
        $Choices,

        [Parameter(Position = 8, ParameterSetName = 'Xml')]
        [ValidateSet('sum', 'count', 'average', 'min', 'max', 'merge', 'plaintext', 'first', 'last')]
        [System.String]
        $Aggregation,

        [Parameter(ParameterSetName = 'Default')]
        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $AddToDefaultView,

        [Parameter(ParameterSetName = 'Default')]
        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $AddToAllContentTypes,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $AllowDeletion,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $AllowHyperlink,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $AllowMultiVote,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $AppendOnly,

        [Parameter(Position = 9, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $AuthoringInfo,

        [Parameter(Position = 10, ParameterSetName = 'Xml')]
        [ValidateSet('Integer', 'Text')]
        [System.String]
        $BaseType,

        [Parameter(Position = 11, ParameterSetName = 'Xml')]
        [ValidateRange(0, 16)]
        [ValidateScript({ $_ -ne 3 -and $_ -ne 13 })]
        [System.Int16]
        $CalType,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $CanToggleHidden,

        [Parameter(Position = 12, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $ClassInfo,

        [Parameter(Position = 13, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $ColName,

        [Parameter(Position = 14, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $ColName2,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $Commas,

        [Parameter(Position = 15, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Customization,

        [Parameter(Position = 16, ParameterSetName = 'Xml')]
        [ValidateRange(0, 7)]
        [System.Int16]
        $Decimals,

        [Parameter(Position = 17, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Description,

        [Parameter(Position = 18, ParameterSetName = 'Xml')]
        [ValidateSet('LTR', 'RTL', 'none')]
        [System.String]
        $Dir,

        [Parameter(Position = 19, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Direction,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $DisplaceOnUpgrade,

        [Parameter(Position = 20, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $DisplayImage,

        [Parameter(Position = 21, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $DisplayNameSrcField,

        [Parameter(Position = 22, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.Int16]
        $DisplaySize,

        [Parameter(Position = 23, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.Int16]
        $Div,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $EnableLookup,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $EnforceUniqueValues,

        [Parameter(Position = 24, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $ExceptionImage,

        [Parameter(Position = 25, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $FieldRef,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $FillInChoice,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $Filterable,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $FilterableNoRecurrence,

        [Parameter(Position = 26, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $ForcedDisplay,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $ForcePromoteDemote,

        [Parameter(Position = 27, ParameterSetName = 'Xml')]
        [ValidateSet('DateOnly', 'DateTime', 'ISO8601', 'ISO8601Basic', 'Dropdown', 'RadioButtons', 'Hyperlink', 'Image')]
        [System.String]
        $Format,

        [Parameter(Position = 28, ParameterSetName = 'Default')]
        [Parameter(Position = 28, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Group,

        [Parameter(Position = 29, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $HeaderImage,

        [Parameter(Position = 30, ParameterSetName = 'Xml')]
        [ValidateRange(0, 1080)]
        [System.Int16]
        $Height,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $Hidden,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $HTMLEncode,

        [Parameter(Position = 31, ParameterSetName = 'Xml')]
        [ValidateSet('active', 'inactive', '')]
        [AllowEmptyString()]
        [System.String]
        $IMEMode,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $Indexed,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $IsolateStyles,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $IsRelationship,

        [Parameter(Position = 32, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $JoinColName,

        [Parameter(Position = 33, ParameterSetName = 'Xml')]
        [ValidateSet('INNER', 'LEFT OUTER', 'RIGHT OUTER')]
        [System.String]
        $JoinType,

        [Parameter(Position = 34, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.Int16]
        $LCID,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $LinkToItem,

        [Parameter(Position = 35, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $LinkToItemAllowed,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $ListItemMenu,

        [Parameter(Position = 36, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $ListItemMenuAllowed,

        [Parameter(Position = 37, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.Double]
        $Max,

        [Parameter(Position = 38, ParameterSetName = 'Xml')]
        [ValidateRange(0, 5000)]
        [System.Int16]
        $MaxLength,

        [Parameter(Position = 39, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.Double]
        $Min,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $Mult,

        [Parameter(Position = 40, ParameterSetName = 'Xml')]
        [ValidateSet('MinusSign', 'Parens')]
        [System.String]
        $NegativeFormat,

        [Parameter(Position = 41, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Node,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $NoEditFormBreak,

        [Parameter(Position = 42, ParameterSetName = 'Xml')]
        [ValidateRange(0, 500)]
        [System.Int16]
        $NumLines,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $Percentage,

        [Parameter(Position = 43, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $PIAttribute,

        [Parameter(Position = 44, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $PITarget,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $PrependId,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $Presence,

        [Parameter(Position = 45, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $PrimaryPIAttribute,

        [Parameter(Position = 46, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $PrimaryPITarget,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $ReadOnly,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $ReadOnlyEnforced,

        [Parameter(Position = 47, ParameterSetName = 'Xml')]
        [ValidateSet('Cascade', 'Restrict', 'None')]
        [System.String]
        $RelationshipDeleteBehavior,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $RenderXMLUsingPattern,

        [Parameter(ParameterSetName = 'Default')]
        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $Required,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $RestrictedMode,

        [Parameter(Position = 48, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $ResultType,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $RichText,

        [Parameter(Position = 49, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.Int16]
        $RowOrdinal,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $ShowAlways,

        [Parameter(Position = 50, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $ShowField,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $ShowInDisplayForm,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $ShowInEditForm,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $ShowInFileDlg,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $ShowInListSettings,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $ShowInNewForm,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $ShowInVersionHistory,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $ShowInViewForms,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $Sortable,

        [Parameter(Position = 51, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $SourceID,

        [Parameter(Position = 52, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $StaticName,

        [Parameter(Position = 53, ParameterSetName = 'Xml')]
        [ValidateNotNull()]
        [AllowEmptyString()]
        [System.String]
        $StorageTZ,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $StripWS,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $SuppressNameDisplay,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $TextOnly,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $UnlimitedLengthInDocumentLibrary,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $URLEncode,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $URLEncodeAsURL,

        [Parameter(Position = 55, ParameterSetName = 'Xml')]
        [ValidateSet('0', '1')]
        [System.String]
        $UserSelectionMode,

        [Parameter(Position = 56, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.Int16]
        $UserSelectionScope,

        [Parameter(Position = 57, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Version,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $Viewable,

        [Parameter(Position = 58, ParameterSetName = 'Xml')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $WebId,

        [Parameter(Position = 59, ParameterSetName = 'Xml')]
        [ValidateRange(0, 1920)]
        [System.Int16]
        $Width,

        [Parameter(ParameterSetName = 'Xml')]
        [Switch]
        $WikiLinking,

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

    if (!($Type)) {
        if (!($suppress)) {
            $Type = $script:Default.FieldTypes | Out-GridView -Title 'Field Type' -OutputMode Single
        }
        else {
            return 1
        }
    }

    Switch ($PSCmdlet.ParameterSetName) {
        'Default' {
            if (!($InternalName)) {
                if (!($suppress)) {
                    $InternalName = Read-Host -Prompt 'InternalName'
                }
                else {
                    return 1
                }
            }
            if (!($DisplayName)) {
                if (!($suppress)) {
                    $DisplayName = Read-Host -Prompt 'DisplayName'
                }
                else {
                    return 1
                }
            }
            $params = [System.Collections.Hashtable]::New()
            $params.Connection = $script:Config.Connection
            $params.List = $List
            $params.InternalName = $InternalName
            $params.DisplayName = $DisplayName
            $params.Type = $Type
            $params.Id = $Id
            if ($Formula) { $params.Formula = $Formula }
            if ($AddToDefaultView) { $params.AddToDefaultView = $true }
            if ($Required) { $params.Required = $true }
            if ($Group) { $params.Group = $Group }
            if ($AddToAllContentTypes) { $params.AddToAllContentTypes = $true }
            if ($Type -in @('Choice', 'MultiChoice')) {
                if ($Choices) { $params.Choices = $Choices }
            }
            
            if (!($suppress)) {
                Write-Host "Creating $DisplayName..." -NoNewline
            }
            Add-PnPField @params
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
            if (!($Name)) {
                if (!($suppress)) {
                    $Name = Read-Host -Prompt 'Name'
                }
                else {
                    return 1
                }
            }
            if (!($Title)) {
                if (!($suppress)) {
                    $Title = Read-Host -Prompt 'Title'
                }
                else {
                    return 1
                }
            }
            $fieldXml = New-Object -TypeName System.Text.StringBuilder
            [void]$fieldXml.Append("<Field Type='$Type' Name='$Name' ID='$Id' Title='$Title'")
            ForEach ($key in $($PSBoundParameters.Keys | Where-Object { $_ -notin @('Site', 'List', 'Type', 'Name', 'Title', 'Default', 'Choices', 'FieldRef', 'Quiet') })) {
                [void]$fieldXml.Append(" ${key}='$($PSBoundParameters[$key])")
            }
            [void]$fieldXml.Append('>')
            if ($Choices) {
                [void]$fieldXml.Append('<CHOICES>')
                ForEach ($choice in $Choices) {
                    [void]$fieldXml.Append("<CHOICE>${choice}</CHOICE>")
                }
                [void]$fieldXml.Append('</CHOICES>')
            }
            if ($Default) {
                [void]$fieldXml.Append("<Default>${Default}</Default>")
            }
            if ($Formula) {
                [void]$fieldXml.Append("<Formula>${Formula}</Formula>")
            }
            if ($FieldRef) {
                [void]$fieldXml.Append('<FieldRefs>')
                ForEach ($ref in $PSBoundParameters['FieldRef']) {
                    [void]$fieldXml.Append("<FieldRef Name='${ref}'/>")
                }
                [void]$fieldXml.Append('</FieldRefs>')
            }

            [void]$fieldXml.Append('</Field>')
            
            if (!($suppress)) {
                Write-Host "Creating $DisplayName..." -NoNewline
            }
            Add-PnPFieldFromXml -List $List -FieldXml $fieldXml.ToString() -Connection $script:Config.Connection
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
