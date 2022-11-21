# Patching.SharePoint PS Module

#region Classes
################################################################################
#                                                                              #
#                                 CLASSES                                      #
#                                                                              #
################################################################################
# . "$PSScriptRoot\$(Split-Path -Path $(Split-Path -Path $PSScriptRoot -Parent) -Leaf).Classes.ps1"
#endregion

#region Variables
################################################################################
#                                                                              #
#                               VARIABLES                                      #
#                                                                              #
################################################################################
try {
    $script:Config = Import-Clixml -Path "$PSScriptRoot\config.xml"
}
catch {
    $script:Config = [ordered]@{
        Url        = $null
        Files      = $null
        Lists      = $null
        Fields     = $null
        Views      = $null
        Connection = $null
    }
    $script:Config | Export-Clixml -Path "$PSScriptRoot\config.xml" -Depth 100
}
finally {
    $script:Default = [ordered]@{
        FieldTypes  = @(
            'Invalid'
            'Integer'
            'Text'
            'Note'
            'DateTime'
            'Counter'
            'Choice'
            'Lookup'
            'Boolean'
            'Number'
            'Currency'
            'URL'
            'Computed'
            'Threading'
            'Guid'
            'MultiChoice'
            'GridChoice'
            'Calculated'
            'File'
            'Attachments'
            'User'
            'Recurrence'
            'CrossProjectLink'
            'ModStat'
            'Error'
            'ContentTypeId'
            'PageSeparator'
            'ThreadIndex'
            'WorkflowStatus'
            'AllDayEvent'
            'WorkflowEventType'
            'Geolocation'
            'OutcomeChoice'
            'Location'
            'Thumbnail'
            'MaxItems'
        )
        QueryElements = @(
            'And'
            'BeginsWith'
            'Contains'
            'DateRangesOverlap'
            'Eq'
            'Geq'
            'Gt'
            'In'
            'Includes'
            'IsNotNull'
            'IsNull'
            'Leq'
            'Lt'
            'Membership'
            'Neq'
            'NotIncludes'
            'Or'
        )
        PageTypes = @(
            'DEFAULTVIEW'
            'DIALOGVIEW'
            'DISPLAYFORM'
            'DISPLAYFORMDIALOG'
            'EDITFORM'
            'EDITFORMDIALOG'
            'NEWFORM'
            'NEWFORMDIALOG'
            'NORMALVIEW'
            'SOLUTIONFORM'
            'VIEW'
        )
        Templates   = @(
            'NoListTemplate'
            'GenericList'
            'DocumentLibrary'
            'Survey'
            'Links'
            'Announcements'
            'Contacts'
            'Events'
            'Tasks'
            'DiscussionBoard'
            'PictureLibrary'
            'DataSources'
            'WebTemplateCatalog'
            'UserInformation'
            'WebPartCatalog'
            'ListTemplateCatalog'
            'XMLForm'
            'MasterPageCatalog'
            'NoCodeWorkflows'
            'WorkflowProcess'
            'WebPageLibrary'
            'CustomGrid'
            'SolutionCatalog'
            'NoCodePublic'
            'ThemeCatalog'
            'DesignCatalog'
            'AppDataCatalog'
            'AppFilesCatalog'
            'DataConnectionLibrary'
            'WorkflowHistory'
            'GanttTasks'
            'HelpLibrary'
            'AccessRequest'
            'PromotedLinks'
            'TasksWithTimelineAndHierarchy'
            'MaintenanceLogs'
            'Meetings'
            'Agenda'
            'MeetingUser'
            'Decision'
            'MeetingObjective'
            'TextBox'
            'ThingsToBring'
            'HomePageLibrary'
            'Posts'
            'Comments'
            'Categories'
            'Facility'
            'Whereabouts'
            'CallTrack'
            'Circulation'
            'Timecard'
            'Holidays'
            'IMEDic'
            'ExternalList'
            'MySiteDocumentLibrary'
            'IssueTracking'
            'AdminTasks'
            'HealthRules'
            'HealthReports'
            'DeveloperSiteDraftApps'
            'ContentCenterModelLibrary'
            'ContentCenterPrimeLibrary'
            'ContentCenterSampleLibrary'
            'AccessApp'
            'AlchemyMobileForm'
            'AlchemyApprovalWorkflow'
            'SharingLinks'
            'HashtagStore'
            'RecipesTable'
            'FormulasTable'
            'WebTemplateExtensionsList'
            'ItemReferenceCollection'
            'ItemReferenceReference'
            'ItemReferenceReferenceCollection'
            'InvalidType'
        )
    }
}
#endregion

#region DotSourcedScripts
################################################################################
#                                                                              #
#                           DOT-SOURCED SCRIPTS                                #
#                                                                              #
################################################################################
. "$PSScriptRoot\Connect-SharePoint.ps1"
. "$PSScriptRoot\Get-SharePointField.ps1"
. "$PSScriptRoot\Get-SharePointList.ps1"
. "$PSScriptRoot\Get-SharePointListItem.ps1"
. "$PSScriptRoot\Get-SharePointUser.ps1"
. "$PSScriptRoot\Get-SharePointView.ps1"
. "$PSScriptRoot\New-SharePointField.ps1"
. "$PSScriptRoot\New-SharePointList.ps1"
. "$PSScriptRoot\New-SharePointListItem.ps1"
. "$PSScriptRoot\New-SharePointView.ps1"
. "$PSScriptRoot\New-SharePointViewQuery.ps1"
. "$PSScriptRoot\Remove-SharePointItem.ps1"
. "$PSScriptRoot\Set-SharePointField.ps1"
. "$PSScriptRoot\Set-SharePointList.ps1"
. "$PSScriptRoot\Set-SharePointListItem.ps1"
. "$PSScriptRoot\Set-SharePointView.ps1"
#endregion

#region ModuleMembers
################################################################################
#                                                                              #
#                              MODULE MEMBERS                                  #
#                                                                              #
################################################################################
Export-ModuleMember -Function Connect-SharePoint
Export-ModuleMember -Function Get-SharePointField
Export-ModuleMember -Function Get-SharePointList
Export-ModuleMember -Function Get-SharePointListItem
Export-ModuleMember -Function Get-SharePointUser
Export-ModuleMember -Function Get-SharePointView
Export-ModuleMember -Function New-SharePointField
Export-ModuleMember -Function New-SharePointList
Export-ModuleMember -Function New-SharePointListItem
Export-ModuleMember -Function New-SharePointView
Export-ModuleMember -Function New-SharePointViewQuery
Export-ModuleMember -Function Remove-SharePointItem
Export-ModuleMember -Function Set-SharePointField
Export-ModuleMember -Function Set-SharePointList
Export-ModuleMember -Function Set-SharePointListItem
Export-ModuleMember -Function Set-SharePointView
#endregion
# SIG # Begin signature block
# MIIFYQYJKoZIhvcNAQcCoIIFUjCCBU4CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUOXmsnRzrHCNH7bg+6Jj425dj
# xbqgggMAMIIC/DCCAeSgAwIBAgIQFnL4oVNG56NIRjNfzwNXejANBgkqhkiG9w0B
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
# DAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUBosHHea/pr3E5JS5ABkud/jf
# xtIwDQYJKoZIhvcNAQEBBQAEggEAUDsD03VV8WV6IyVB0lpgn5QJ3KF1nRD5QlwU
# c0neirCwC6siM7QSZInLUrkYA8qe9RKwoi36kbm7FcDeBYuRQ9FaIWet7JZjfI+P
# cNGWbcQbqlnm6vr+HD1OOA45wiSQ7ddmfyNWOTqjjFrw+vyhrP4RaNfupiJAYLVV
# iTy/zR/rlc8mpTvEU9hKFTkhJm1KduvltgQlmMAZbp5egUNDrWunN+CX2D1TLIpA
# 3HAbbhTSVLcI27/9pOE9bNs+7D/+7JjfJrD28Jrl9Jl/xT2H0U5Fgyf/BAFLk3jy
# 6xX+F1l/ddaI5dheiSMAyH3HDu37hwdtRlj+LQUnBHYvxBFcZA==
# SIG # End signature block
