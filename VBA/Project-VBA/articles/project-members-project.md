---
title: Project Members (Project)
ms.prod: PROJECTSERVER
ms.assetid: 0d67dded-e3ed-7dff-b412-0cd7a3627c5c
---


# Project Members (Project)
Represents one project in the set of open projects. The  **Project** object is a member of the **[Projects](projects-object-project.md)** collection.

Represents one project in the set of open projects. The  **Project** object is a member of the **[Projects](projects-object-project.md)** collection.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[Activate](project-activate-event-project.md)|Occurs when switching to the project from another project, including when the project is opened or created.|
|[BeforeClose](project-beforeclose-event-project.md)|Occurs before a project is closed.|
|[BeforePrint](project-beforeprint-event-project.md)|Occurs before a project is printed.|
|[BeforeSave](project-beforesave-event-project.md)|Occurs before a project is saved.|
|[Calculate](project-calculate-event-project.md)|Occurs when a project schedule is recalculated.|
|[Change](project-change-event-project.md)|Occurs when a change is made to data in the project. An action affecting several items at once is considered to be one change.|
|[Deactivate](project-deactivate-event-project.md)|Occurs when switching from the current project to another project.|
|[Open](project-open-event-project.md)|Occurs when the project opens, but before the  **Activate** event.|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Activate](project-activate-method-project.md)|Activates the project.|
|[AppendNotes](project-appendnotes-method-project.md)|Appends text to the Notes field.|
|[CheckIn](project-checkin-method-project.md)|Checks in the working copy of the project from a local computer to the SharePoint document library, and sets the local project to read-only so that it cannot be edited locally.|
|[CheckoutProject](project-checkoutproject-method-project.md)|Checks out an open project that is currently in read-only mode.|
|[DeliverableAcceptChanges](project-deliverableacceptchanges-method-project.md)|Accepts the changes on the server for a deliverable.|
|[DeliverableCreate](project-deliverablecreate-method-project.md)|Creates a deliverable for a published project that has a project workspace.|
|[DeliverableDelete](project-deliverabledelete-method-project.md)|Deletes the deliverable.|
|[DeliverableDependencyCreate](project-deliverabledependencycreate-method-project.md)|Creates a dependency on a deliverable and links the dependency to a task in the project.|
|[DeliverableDependencyDelete](project-deliverabledependencydelete-method-project.md)|Deletes the dependency on the deliverable.|
|[DeliverableLinkToProject](project-deliverablelinktoproject-method-project.md)|Links a deliverable or a dependency to a project.|
|[DeliverableLinkToTask](project-deliverablelinktotask-method-project.md)|Links a deliverable to a task.|
|[DeliverableRefreshServerCache](project-deliverablerefreshservercache-method-project.md)|Checks for updates on the server and refreshes the cache for deliverable and dependencies for the project.|
|[DeliverablesClearAll](project-deliverablesclearall-method-project.md)|Clears all deliverables in the project.|
|[DeliverablesGetByProject](project-deliverablesgetbyproject-method-project.md)|Gets a list of all deliverables for the specified enterprise project in the XML member of the returned object. Project Professional only.|
|[DeliverablesGetProviderProjects](project-deliverablesgetproviderprojects-method-project.md)|Returns a list of all of the projects that have deliverables.|
|[DeliverablesGetServerCachedXml](project-deliverablesgetservercachedxml-method-project.md)|Gets the cached XML data from Project Server for all of the deliverables and dependencies in a project.|
|[DeliverablesGetXml](project-deliverablesgetxml-method-project.md)|Gets the XML data from Project Professional for all of the deliverables and dependencies in a project.|
|[DeliverableUpdate](project-deliverableupdate-method-project.md)|Updates the properties of a deliverable.|
|[ExportAsFixedFormat](project-exportasfixedformat-method-project.md)|Exports the active project as a document in a custom PDF or XPS format.|
|[GetDisplayNameFromObjectMatchingID](project-getdisplaynamefromobjectmatchingid-method-project.md)|Returns the display name of an object.|
|[GetObjectMatchingID](project-getobjectmatchingid-method-project.md)|Returns the matching identification name of an object.|
|[GetServerProjectGuid](project-getserverprojectguid-method-project.md)|Returns the GUID for the enterprise project.|
|[GetTaskIndexByGuid](project-gettaskindexbyguid-method-project.md)|Returns the local task identification number (ID) for the specified task.|
|[GetWinprojURLs](project-getwinprojurls-method-project.md)|Returns the various URLs associated with the active enterprise project as an XML string.|
|[HideCheckoutMsgBar](project-hidecheckoutmsgbar-method-project.md)|Hides the project checkout message bar.|
|[ImportResourceErrorCount](project-importresourceerrorcount-method-project.md)|Returns the number of errors generated by a resource import operation.|
|[LevelClearDates](project-levelcleardates-method-project.md)|Sets the leveling range to include the entire project.|
|[LocalResourceCount](project-localresourcecount-method-project.md)|Returns the number of local resources in the project.|
|[LocalResourceErrorCount](project-localresourceerrorcount-method-project.md)|Returns the number of local resource errors.|
|[MakeServerURLTrusted](project-makeserverurltrusted-method-project.md)|Adds the URL specified in the  **[ServerURL](a204c795-73a3-4ce2-a582-3afd951914c7.md)** property to the **Trusted sites** zone in the **Security** tab of the **Internet Options** dialog box in Internet Explorer.|
|[ReadWssData](project-readwssdata-method-project.md)|Returns the Project Workspace URLs for the active enterprise project as an XML string.|
|[ResourceCount](project-resourcecount-method-project.md)|Returns the total number of resources in the project, including both local and enterprise resources.|
|[ResourceErrorCount](project-resourceerrorcount-method-project.md)|Returns the number of resource errors.|
|[SaveAs](project-saveas-method-project.md)|Saves a file that is not the active project under a new file name.|
|[SetCustomUI](project-setcustomui-method-project.md)|Sets the internal XML value for a custom ribbon user interface of the project.|
|[SetObjectMatchingID](project-setobjectmatchingid-method-project.md)|Sets the matching identification value of an object in the  **Organizer** dialog box, for example to change the view specified by "Gantt Chart".|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[AcceptNewExternalData](project-acceptnewexternaldata-property-project.md)|**True** if new or changed data relating to an external task is automatically accepted when the project is opened. Read/write **Boolean**.|
|[AdministrativeProject](project-administrativeproject-property-project.md)|In Microsoft Project 2003, indicates whether the project is an administrative project. Read/write  **Boolean**. Not used in later versions of Project.|
|[AllowTaskDelegation](project-allowtaskdelegation-property-project.md)|**True** if Project Server users can delegate tasks to other resources in the project. Read/write **Boolean**. .|
|[AndMoveCompleted](project-andmovecompleted-property-project.md)|**True** if the actual, completed portion of a task that is scheduled before the status date is moved to end at the status date. Read/write **Boolean**.|
|[AndMoveRemaining](project-andmoveremaining-property-project.md)|**True** if the remaining work on a task that is scheduled after the status date is moved to start at the status date. Read/write **Boolean**.|
|[Application](project-application-property-project.md)|Gets the  **[Application](application-object-project.md)** object. Read-only **Application**.|
|[AskForCompletedWork](project-askforcompletedwork-property-project.md)|Gets or sets the way completed work is reported in team status messages. Read/write  **PjTeamStatusCompletedWork**.|
|[AutoAddResources](project-autoaddresources-property-project.md)|**True** if new resources are automatically created as they are assigned. **False** if Project prompts before creating new resources. Read/write **Boolean**.|
|[AutoCalcCosts](project-autocalccosts-property-project.md)|**True** if Project always calculates actual costs. **False** if users can enter actual costs, and Project does not calculate actual costs. Read/write **Boolean**.|
|[AutoFilter](project-autofilter-property-project.md)|Gets or sets whether the AutoFilter feature is turned on for a project. Read/write  **Boolean**.|
|[AutoLinkTasks](project-autolinktasks-property-project.md)|**True** if Project automatically links sequential tasks when you cut, move, or insert tasks. Read/write **Boolean**.|
|[AutoSplitTasks](project-autosplittasks-property-project.md)|**True** if Project automatically splits tasks into parts for work complete and work remaining. Read/write **Boolean**.|
|[AutoTrack](project-autotrack-property-project.md)|**True** if Project automatically updates the work and costs of resources assigned to a task when the percent complete changes. Read/write **Boolean**.|
|[BaseCalendars](project-basecalendars-property-project.md)|Gets a  **[Calendars](calendar-object-project.md)** collection representing all base calendars in the active project. Read-only **Calendars**.|
|[BaselineSavedDate](project-baselinesaveddate-property-project.md)|Gets date the specified baseline was last saved. Read-only  **Variant**.|
|[BuiltinDocumentProperties](project-builtindocumentproperties-property-project.md)|Gets a  **DocumentProperties** collection representing the built-in properties of the document. Read-only **Object**.|
|[Calendar](project-calendar-property-project.md)|Gets a  **[Calendar](calendar-object-project.md)** object representing a calendar for the project. Read-only **Calendar**.|
|[CanCheckIn](project-cancheckin-property-project.md)|**True** if Project Professional can check in a project to Project Server. Read-only **Boolean**.|
|[CodeName](project-codename-property-project.md)|Gets the code name for the project. Read-only  **String**.|
|[CommandBars](project-commandbars-property-project.md)|Gets a  **CommandBars** collection that represents all the command bars in the project. Read-only **CommandBars**.|
|[Container](project-container-property-project.md)|Gets the object that contains the embedded project. Read-only  **Object**.|
|[CreationDate](project-creationdate-property-project.md)|Gets the date a project was created. Read-only  **Variant**.|
|[CurrencyCode](project-currencycode-property-project.md)|Project property for the three-character ISO standard currency code of the project. Read/write  **String**.|
|[CurrencyDigits](project-currencydigits-property-project.md)|Sets or returns the number of digits following the decimal separator character in currency values. Read/write  **Integer**.|
|[CurrencySymbol](project-currencysymbol-property-project.md)|Gets or sets the characters that denote currency values. Read/write  **String**.|
|[CurrencySymbolPosition](project-currencysymbolposition-property-project.md)|Gets or sets the location of the currency symbol. Read/write  **PjPlacement**.|
|[CurrentDate](project-currentdate-property-project.md)|Gets or sets the current date for a project. Read/write  **Variant**.|
|[CurrentFilter](project-currentfilter-property-project.md)|Gets the name of the active filter for a project. Read-only  **String**.|
|[CurrentGroup](project-currentgroup-property-project.md)|Gets the name of the active group for the active project. Read-only  **String**.|
|[CurrentTable](project-currenttable-property-project.md)|Gets the name of the active table for a project. Read-only  **String**.|
|[CurrentView](project-currentview-property-project.md)|Gets the name of the active view for a project. Read-only  **String**.|
|[CustomDocumentProperties](project-customdocumentproperties-property-project.md)|Gets a  **DocumentProperties** collection representing the custom properties of the document. Read-only **Object**.|
|[DatabaseProjectUniqueID](project-databaseprojectuniqueid-property-project.md)|Gets the project unique ID for a project stored in a database. Read/write  **Variant**.|
|[DayLabelDisplay](project-daylabeldisplay-property-project.md)|Gets or sets the abbreviation for "day" that is displayed for values such as durations, delays, slack, and work. Read/write  **Integer**.|
|[DaysPerMonth](project-dayspermonth-property-project.md)|Gets or sets the number of days per month for tasks in a project. Read/write  **Double**.|
|[DefaultDurationUnits](project-defaultdurationunits-property-project.md)|Gets or sets the default duration units. Read/write  **PjUnit**.|
|[DefaultEarnedValueMethod](project-defaultearnedvaluemethod-property-project.md)|Gets or sets the default method for calculating earned value for a project. Read/write  **PjEarnedValueMethod**.|
|[DefaultEffortDriven](project-defaulteffortdriven-property-project.md)|**True** if new tasks are effort-driven by default. Read/write **Boolean**.|
|[DefaultFinishTime](project-defaultfinishtime-property-project.md)|Gets or sets the default finish time of the project. Read/write  **Variant**.|
|[DefaultFixedCostAccrual](project-defaultfixedcostaccrual-property-project.md)|Gets or sets the default method used to accrue fixed task costs in the project. Read/write  **PjAccrueAt**.|
|[DefaultResourceOvertimeRate](project-defaultresourceovertimerate-property-project.md)|Gets or sets the default overtime rate of pay for resources. Read/write  **Variant**.|
|[DefaultResourceStandardRate](project-defaultresourcestandardrate-property-project.md)|Gets or sets the default standard rate of pay for resources. Read/write  **Variant**.|
|[DefaultStartTime](project-defaultstarttime-property-project.md)|Gets or sets the default start time for the project. Read/write  **Variant**.|
|[DefaultTaskType](project-defaulttasktype-property-project.md)|Gets or sets the default task type. Read/write  **PjTaskFixedType**.|
|[DefaultWorkUnits](project-defaultworkunits-property-project.md)|Gets or sets the default work units for the project. Read/write  **PjUnit**.|
|[DetectCycle](project-detectcycle-property-project.md)|Gets a  **Tasks** collection that contains a set of circular task dependencies, if circular task references exist. Read-only **Tasks**.|
|[DisplayProjectSummaryTask](project-displayprojectsummarytask-property-project.md)|**True** if the summary task for a project is visible. Read/write **Boolean**.|
|[DocumentLibraryVersions](project-documentlibraryversions-property-project.md)|Gets a  **DocumentLibraryVersions** collection for the specified project. Read-only **DocumentLibraryVersions**.|
|[EarnedValueBaseline](project-earnedvaluebaseline-property-project.md)|Gets or sets the baseline for the earned values of tasks. Read/write  **PjBaselines**.|
|[Engagements](project-engagements-property-project.md)||
|[EnterpriseActualsSynched](project-enterpriseactualssynched-property-project.md)|**True** if the actual work or actual overtime in a project is synchronized with the actual work or actual overtime that has been submitted and updated from the timesheet system. Read/write **Boolean**.|
|[ExpandDatabaseTimephasedData](project-expanddatabasetimephaseddata-property-project.md)|**True** if timephased data is expanded to a readable format in the database. **False** if timephased data is in a compressed binary format. Read/write **Boolean**.|
|[FollowedHyperlinkColor](project-followedhyperlinkcolor-property-project.md)|Gets or sets the color used to denote followed hyperlinks. Read/write  **PjColor**.|
|[FollowedHyperlinkColorEx](project-followedhyperlinkcolorex-property-project.md)|Gets or sets the color used to denote followed hyperlinks. Read/write  **Long**.|
|[FullName](project-fullname-property-project.md)|Gets the path and file name of a project. Read-only  **String**.|
|[HasPassword](project-haspassword-property-project.md)|**True** if a project has a password. Read-only **Boolean**.|
|[HonorConstraints](project-honorconstraints-property-project.md)|**True** if tasks honor their constraint dates. Read/write **Boolean**.|
|[HourLabelDisplay](project-hourlabeldisplay-property-project.md)|Gets or sets the abbreviation for "hour" that is displayed for values such as durations, delays, slack, and work. Read/write  **Integer**.|
|[HoursPerDay](project-hoursperday-property-project.md)|Gets or sets the number of hours per day for tasks in a project. Read/write  **Double**.|
|[HoursPerWeek](project-hoursperweek-property-project.md)|Gets or sets the number of hours per week for tasks in a project. Read/write  **Double**.|
|[HyperlinkColor](project-hyperlinkcolor-property-project.md)|Gets or sets the color used to denote unfollowed hyperlinks. Read/write  **PjColor**.|
|[HyperlinkColorEx](project-hyperlinkcolorex-property-project.md)|Gets or sets a hexadecimal representation of the color used to denote unfollowed hyperlinks. Read/write  **Long**.|
|[ID](project-id-property-project.md)|Gets the identification number of a project. Read-only  **Long**.|
|[Index](project-index-property-project.md)|Gets the index of a  **Project** object in the containing **Projects** collection. Read-only **Variant**.|
|[IsCheckoutMsgBarVisible](project-ischeckoutmsgbarvisible-property-project.md)|Gets whether the checkout message bar is visible. Read-only  **Boolean**.|
|[IsCheckoutOSVisible](project-ischeckoutosvisible-property-project.md)|Gets whether the  **Check Out** button is visible in the Backstage view. Read-only **Boolean**.|
|[IsTemplate](project-istemplate-property-project.md)|Gets or sets a value that indicates whether the project is a template. Read/write  **Boolean**.|
|[KeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled](project-keeptaskonnearestworkingtimewhenmadeautoscheduled-property-project.md)|**True** if task scheduling respects the current calendar when a task is converted from manual to automatic; otherwise, **False**. Read/write **Boolean**.|
|[LastPrintedDate](project-lastprinteddate-property-project.md)|Gets the date a project was last printed. Read-only  **Variant**.|
|[LastSaveDate](project-lastsavedate-property-project.md)|Gets the date a project was last saved. Read-only  **Variant**.|
|[LastSavedBy](project-lastsavedby-property-project.md)|Gets the name of the user who last saved a project. Read-only  **String**.|
|[LastWssSyncDate](project-lastwsssyncdate-property-project.md)||
|[LevelEntireProject](project-levelentireproject-property-project.md)|**True** if all resources in the project are leveled. **False** if only overallocated resources within specified dates are leveled. Read/write **Boolean**.|
|[LevelFromDate](project-levelfromdate-property-project.md)|Gets or sets the starting date of a range in which overallocated resources are leveled. The default is the project start date or the last entered date value. Read/write  **Variant**.|
|[LevelToDate](project-leveltodate-property-project.md)|Gets or sets the ending date of a range in which overallocated resources are leveled. The default is the project finish date or the last entered date value. Read/write  **Variant**.|
|[ManuallyScheduledTasksAutoRespectLinks](project-manuallyscheduledtasksautorespectlinks-property-project.md)|**True** if predecessor and successor task links are maintained when a task is converted from automatic to manual; otherwise, **False**. Read/write **Boolean**.|
|[MapList](project-maplist-property-project.md)|Gets a  **[List](list-object-project.md)** object representing the list of data maps in the project. Read-only **List**.|
|[MinuteLabelDisplay](project-minutelabeldisplay-property-project.md)|Gets or sets the abbreviation for "minute" that is displayed for values such as durations, delays, slack, and work. Read/write  **Integer**.|
|[MonthLabelDisplay](project-monthlabeldisplay-property-project.md)|Gets or sets the abbreviation for "month" that is displayed for values such as durations, delays, slack, and work. Read/write  **Integer**.|
|[MoveCompleted](project-movecompleted-property-project.md)|**True** if a task that is scheduled after the status date has actual progress entered against it and the actual, completed portion of the task is moved so the completed work ends on the status date. Read/Write **Boolean**.|
|[MoveRemaining](project-moveremaining-property-project.md)|**True** if the remaining portion of a task that is scheduled before the status date is moved to start at the status date. Read/Write **Boolean**.|
|[MultipleCriticalPaths](project-multiplecriticalpaths-property-project.md)|**True** if Project calculates multiple critical paths for the project. **False** if only one critical path is calculated. Read/write **Boolean**.|
|[Name](project-name-property-project.md)|Gets the name of a  **Project** object. Read-only **String**.|
|[NewTasksCreatedAsManual](project-newtaskscreatedasmanual-property-project.md)|**True** if new tasks are created as manually scheduled tasks. **False** if new tasks are automatically scheduled. Read/write **Boolean**.|
|[NewTasksEstimated](project-newtasksestimated-property-project.md)|**True** if new tasks in the active project have estimated durations. Read/write **Boolean**.|
|[NumberOfResources](project-numberofresources-property-project.md)|Gets the number of resources in a project, not including blank entries. Read-only  **Long**.|
|[NumberOfTasks](project-numberoftasks-property-project.md)|Gets the number of tasks in a project, not including blank entries. Read-only  **Long**.|
|[OutlineChildren](project-outlinechildren-property-project.md)|Gets a  **[Tasks](task-object-project.md)** collection representing the children of a task in the outline structure. Read-only **Tasks**.|
|[OutlineCodes](project-outlinecodes-property-project.md)|Gets an **[OutlineCodes](outlinecodes-object-project.md)** collection of all outline codes defined for resources and tasks in the project. Read-only **OutlineCodes**.|
|[Parent](project-parent-property-project.md)|Gets the parent of the  **Project** object. Read-only **Object**.|
|[Path](project-path-property-project.md)|Gets the path of the open project. Read-only  **String**.|
|[PhoneticType](project-phonetictype-property-project.md)|Gets or sets the type of characters used to display phonetic information. Read/write  **PjPhoneticType**.|
|[ProjectFinish](project-projectfinish-property-project.md)|Gets or sets the finish date for a project. Read/write  **Variant**.|
|[ProjectGuideContent](project-projectguidecontent-property-project.md)|Gets or sets the name of the XML schema being used by the Project Guide. Read/write  **String**.|
|[ProjectGuideFunctionalLayoutPage](project-projectguidefunctionallayoutpage-property-project.md)|Gets or sets the Project Guide functional layout page for the specified project. Read/write  **String**.|
|[ProjectGuideSaveBuffer](project-projectguidesavebuffer-property-project.md)|Gets or sets an XML string representing the save buffer of the Project Guide. Read/write  **String**.|
|[ProjectGuideUseDefaultContent](project-projectguideusedefaultcontent-property-project.md)|**True** if the Project Guide uses the default content. **False** if you want to use custom content for the Project Guide. Read/write **Boolean**.|
|[ProjectGuideUseDefaultFunctionalLayoutPage](project-projectguideusedefaultfunctionallayoutpage-property-project.md)|**True** if Project uses the default Project Guide. **False** if you are customizing the Project Guide. Read/write **Boolean**.|
|[ProjectNamePrefix](project-projectnameprefix-property-project.md)|Gets the prefix of the project name of the specified project. Read-only  **String**.|
|[ProjectNotes](project-projectnotes-property-project.md)|Gets or sets the notes for the project. Read/write  **String**.|
|[ProjectServerUsedForTracking](project-projectserverusedfortracking-property-project.md)|**True** if Project Server is used for tracking the specified project. Read/write **Boolean**.|
|[ProjectStart](project-projectstart-property-project.md)|Gets or sets the start date for a project. Read/write  **Variant**.|
|[ProjectSummaryTask](project-projectsummarytask-property-project.md)|Gets a  **[Task](task-object-project.md)** object representing the project summary task for the active project. Read-only **Task**.|
|[ReadOnly](project-readonly-property-project.md)|**True** if a project has read-only access. Read-only **Boolean**.|
|[ReadOnlyRecommended](project-readonlyrecommended-property-project.md)|**True** if the project should be opened with read-only access. Read-only **Boolean**.|
|[ReceiveNotifications](project-receivenotifications-property-project.md)|**True** if the manager receives notification when new messages arrive. Obsolete in Project. Read/write **Boolean**.|
|[RemoveFileProperties](project-removefileproperties-property-project.md)|**True** if Project removes user information from revisions and the project **Properties** dialog box upon saving a document. Read/write **Boolean**.|
|[ReportList](project-reportlist-property-project.md)|Deprecated in Project. |
|[Reports](project-reports-property-project.md)|Gets the collection of custom reports in the project. Read-only  **Reports**.|
|[ResourceFilterList](project-resourcefilterlist-property-project.md)|Gets a  **[List](list-object-project.md)** object representing all resource filters in the project. Read-only **List**.|
|[ResourceFilters](project-resourcefilters-property-project.md)|Gets a  **[Filters](filters-object-project.md)** collection that contains the resource filters of the project. Read-only **Filters**.|
|[ResourceGroupList](project-resourcegrouplist-property-project.md)|Gets a  **[List](list-object-project.md)** object representing the resource groups in the active project. Read-only **List**.|
|[ResourceGroups](project-resourcegroups-property-project.md)|Gets a  **[ResourceGroups](resourcegroups-object-project.md)** collection that contains all the resource-based group definitions in the project. Read-only **ResourceGroups**.|
|[ResourceGroups2](project-resourcegroups2-property-project.md)|Gets a  **[ResourceGroups2](resourcegroups2-object-project.md)** collection that represents all of the resource groups based on **Group2** objects. Read-only **ResourceGroups2**.|
|[ResourcePoolName](project-resourcepoolname-property-project.md)|Gets the name of the enterprise resource pool that a project uses in Project Professional. Read-only  **String**.|
|[Resources](project-resources-property-project.md)|Gets a  **[Resources](resources-object-project.md)** collection representing the resources in a **Project**. Read-only **Object**.|
|[ResourceTableList](project-resourcetablelist-property-project.md)|Gets a  **[List](list-object-project.md)** object representing all resource tables in the project. Read-only **List**.|
|[ResourceTables](project-resourcetables-property-project.md)|Gets a  **[Tables](tables-object-project.md)** collection that contains the resource tables of the project. Read-only **Tables**.|
|[ResourceViewList](project-resourceviewlist-property-project.md)|Gets a  **[List](list-object-project.md)** object representing all resource views in the active project. Read-only **List**.|
|[RevisionNumber](project-revisionnumber-property-project.md)|Gets the number of times a project has been saved. Read-only  **String**.|
|[Saved](project-saved-property-project.md)|**True** if a project has not changed since it was last saved. Read-only **Boolean**.|
|[ScheduleFromStart](project-schedulefromstart-property-project.md)|**True** if Project calculates the project schedule forward from the start date. **False** if the schedule is calculated backward from the finish date. Read/write **Boolean**.|
|[SendHyperlinkNote](project-sendhyperlinknote-property-project.md)|**True** if resources receive e-mail notification Obsolete in Project. Read/write **Boolean**.|
|[ServerIdentification](project-serveridentification-property-project.md)|Gets or sets the way Project Professional users are identified to Project Server. Read/write  **PjAuthentication**.|
|[ServerURL](project-serverurl-property-project.md)|Gets the URL of the Project Web App instance with which Project Professional is connected. For a synchronized SharePoint task list, gets or sets an arbitrary value that has no effect on the project. Read/write  **String**.|
|[SharedWorkspace](project-sharedworkspace-property-project.md)|Gets a  **SharedWorkspace** object that represents the document workspace for the project. Read-only **SharedWorkspace**.|
|[ShowCriticalSlack](project-showcriticalslack-property-project.md)|Gets or sets how much slack causes a task to be displayed as a critical task. Read/write  **Long**.|
|[ShowCrossProjectLinksInfo](project-showcrossprojectlinksinfo-property-project.md)|**True** if the **Links between Projects** dialog box appears when a project containing cross-project links is opened. Read-only **Boolean**.|
|[ShowEstimatedDuration](project-showestimatedduration-property-project.md)|**True** if task durations in the project are displayed with the estimated character. Read/write **Boolean**.|
|[ShowExternalPredecessors](project-showexternalpredecessors-property-project.md)|**True** if predecessor tasks linked from an external project should be displayed. Read-only **Boolean**.|
|[ShowExternalSuccessors](project-showexternalsuccessors-property-project.md)|**True** if successor tasks linked from an external project should be displayed. Read-only **Boolean**.|
|[ShowTaskSuggestions](project-showtasksuggestions-property-project.md)|**True** if task suggestions in the active project are displayed; otherwise, **False**. Read/write **Boolean**.|
|[ShowTaskWarnings](project-showtaskwarnings-property-project.md)|**True** if task warnings in the active project are displayed; otherwise, **False**. Read/write **Boolean**.|
|[SpaceBeforeTimeLabels](project-spacebeforetimelabels-property-project.md)|**True** if a time value should be separated from its time label by a space. Read/write **Boolean**.|
|[SpreadCostsToStatusDate](project-spreadcoststostatusdate-property-project.md)|**True** if edits to total actual cost are spread to the status date, or to the current date if the status date is "NA". **False** if edits are spread to the calculated stop date of the task. Read/write **Boolean**.|
|[SpreadPercentCompleteToStatusDate](project-spreadpercentcompletetostatusdate-property-project.md)|**True** if edits to total task percent complete are spread to the status date, or to the current date if the status date is "NA". **False** if edits are spread to the calculated stop date of the task. Read/write **Boolean**.|
|[StartOnCurrentDate](project-startoncurrentdate-property-project.md)|**True** if new tasks start on the current date. **False** if new tasks start on the project start date. Read/write **Boolean**.|
|[StartWeekOn](project-startweekon-property-project.md)|Gets or sets the first day of the week for the project. Read/write  **PjWeekday**.|
|[StartYearIn](project-startyearin-property-project.md)|Gets or sets the month number for the start of the fiscal year for the project. Read/write  **PjMonth**.|
|[StatusDate](project-statusdate-property-project.md)|Gets or sets the current status date for the project. If there is no status date, returns "NA". Read/write  **Variant**.|
|[Subprojects](project-subprojects-property-project.md)|Gets a  **[Subprojects](subproject-object-project.md)** collection representing subprojects in the master project. Read-only **Subprojects**.|
|[TaskErrorCount](project-taskerrorcount-property-project.md)|Gets or sets the number of task errors associated with a project. Read/write  **Long**.|
|[TaskFilterList](project-taskfilterlist-property-project.md)|Gets a  **[List](list-object-project.md)** object representing all task filters in the project. Read-only **List**.|
|[TaskFilters](project-taskfilters-property-project.md)|Gets a  **[Filters](filter-object-project.md)** collection of the task filters in the project. Read-only **Filters**.|
|[TaskGroupList](project-taskgrouplist-property-project.md)|Gets a  **[List](list-object-project.md)** object representing the task groups in the active project. Read-only **List**.|
|[TaskGroups](project-taskgroups-property-project.md)|Gets a  **[TaskGroups](taskgroups-object-project.md)** collection representing all the task-based **Group** definitions in the project. Read-only **TaskGroups**.|
|[TaskGroups2](project-taskgroups2-property-project.md)|Gets a  **[TaskGroups2](taskgroups2-object-project.md)** collection that represents all the task-based **Group2** definitions in the specified project. Read-only **TaskGroups2**.|
|[Tasks](project-tasks-property-project.md)|Gets a  **[Tasks](task-object-project.md)** collection representing the tasks in the project. Read-only **Tasks**.|
|[TaskTableList](project-tasktablelist-property-project.md)|Gets a  **[List](list-object-project.md)** object representing all task tables in the project. Read-only **List**.|
|[TaskTables](project-tasktables-property-project.md)|Gets a  **[Tables](table-object-project.md)** collection representing the task tables in the project. Read-only **Tables**.|
|[TaskViewList](project-taskviewlist-property-project.md)|Gets a  **[List](list-object-project.md)** object representing all task views in the project. Read-only **List**.|
|[Template](project-template-property-project.md)|Gets the name of the template associated with a project. Read-only  **String**.|
|[Timeline](project-timeline-property-project.md)||
|[TrackingMethod](project-trackingmethod-property-project.md)|Gets or sets the tracking method used by Project Server for the project. Read/write  **PjProjectServerTrackingMethod**.|
|[Type](project-type-property-project.md)|Gets the type of a project. Read-only  **PjProjectType**.|
|[UnderlineHyperlinks](project-underlinehyperlinks-property-project.md)|**True** if hyperlinks are underlined. Read/write **Boolean**.|
|[UniqueID](project-uniqueid-property-project.md)|Gets the unique identification number of the project, which is actually the  **UniqueID** value of the project summary task. Read-only **Long**.|
|[UpdateProjOnSave](project-updateprojonsave-property-project.md)|**True** if Project updates the project schedule when saving a file. Obsolete in Project. Read-only **Boolean**.|
|[UseFYStartYear](project-usefystartyear-property-project.md)|**True** if a fiscal year is determined by the year of the first month of that fiscal year. **False** if determined by the last month of the fiscal year. Read/write **Boolean**.|
|[UserControl](project-usercontrol-property-project.md)|**True** if the user directly opens or creates the project. Read-only **Boolean**.|
|[UtilizationDate](project-utilizationdate-property-project.md)||
|[UtilizationType](project-utilizationtype-property-project.md)||
|[VBASigned](project-vbasigned-property-project.md)|**True** if the Microsoft Visual Basic for Applications project is digitally signed. Read/write **Boolean**.|
|[VBProject](project-vbproject-property-project.md)|Gets a  **VBProject** object that represents the Microsoft Visual Basic project. Read-only **VBProject**.|
|[VersionName](project-versionname-property-project.md)|Gets the version name of the project. Obsolete in Project. Read-only  **String**.|
|[ViewList](project-viewlist-property-project.md)|Gets the  **[List](list-object-project.md)** object for the project. Read-only **List**.|
|[Views](project-views-property-project.md)|Gets a  **[Views](view-object-project.md)** collection representing the views of the project. Read-only **Views**.|
|[ViewsCombination](project-viewscombination-property-project.md)|Gets a  **[ViewsCombination](viewcombination-object-project.md)** collection representing the combination views of the project. Read-only **ViewsCombination**.|
|[ViewsSingle](project-viewssingle-property-project.md)|Gets a  **[ViewsSingle](viewsingle-object-project.md)** collection representing the single views of the project. Read-only **ViewsSingle**.|
|[WBSCodeGenerate](project-wbscodegenerate-property-project.md)|**True** if a work breakdown structure (WBS) code is automatically generated for new tasks in the project. Read/write **Boolean**.|
|[WBSVerifyUniqueness](project-wbsverifyuniqueness-property-project.md)|**True** if an edited work breakdown structure (WBS) code is verified to be unique. Read/write **Boolean**.|
|[WeekLabelDisplay](project-weeklabeldisplay-property-project.md)|Gets or sets the abbreviation for "week" that is displayed for values such as durations, delays, slack, and work. Read/write  **Integer**.|
|[Windows](project-windows-property-project.md)|Gets a  **[Windows](windows-object-project.md)** collection representing the open windows in the project. Read-only **Windows**.|
|[Windows2](project-windows2-property-project.md)|Gets a  **[Windows2](windows2-object-project.md)** collection representing the open windows in the project. Read-only **Windows2**.|
|[WriteReserved](project-writereserved-property-project.md)|**True** if a password is required to open a project for read/write access. Read-only **Boolean**.|
|[YearLabelDisplay](project-yearlabeldisplay-property-project.md)|Gets or sets how the year label displays in rates. Read/write  **Integer**.|

