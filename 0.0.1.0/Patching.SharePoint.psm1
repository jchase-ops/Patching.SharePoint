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
