---
title: Application Object (Project)
ms.prod: PROJECTSERVER
api_name:
- Project.Application
ms.assetid: 8eb91712-7784-a102-38c0-19bb056c27e9
---


# Application Object (Project)

Represents the entire Project application. The  **Application** object contains:


- Application-wide settings and options (many of the options in the  **Options** dialog box on the **Tools** menu, for example).
    
- Properties that return top-level objects, such as  **ActiveCell**, **ActiveProject**, and so forth.
    
- Methods that act on application-wide elements, such as views, selections, editing actions, and so forth.
    

## Using the Application Object

Use the  **[Application](http://msdn.microsoft.com/library/project-application-property-project%28Office.15%29.aspx)** property to return an **Application** object in Project . The following example applies the **Windows** property to the **Application** object.


```
Application.Windows("Project1.mpp").Activate
```


## Using Project From Another Application: Late Binding

The following example creates the Microsoft Project  **Application** object at run time, creates a new project, adds a task, saves the project, and then closes the Project . For example, copy and paste the **CreateProject_Late** macro to the **ThisDocument** module in the Visual Basic Editor (VBE) of Word.


 **Note**  Because the application queries the  **MSProject.Application** type library only at run time, Microsoft IntelliSense is not available and performance is relatively poor with late binding. Scripting languages, such as JavaScript and VBScript, require late binding. VBScript supports only the generic **Object** and **Variant** data types. For better performance in VBA and other compiled languages, you should use early binding by setting a reference to the Project type library.


```
Sub CreateProject_Late() 
    Dim pjApp As Object 
    Set pjApp = CreateObject("MSProject.Application") 
    pjApp.Visible = True 
    pjApp.FileNew 
    pjApp.ActiveProject.Tasks.Add "Hang clocks" 
    pjApp.FileSaveAs "Clocks.mpp" 
    pjApp.FileClose 
    pjApp.Quit 
End Sub
```

If you do not set the  **Visible** property to **True**, the Project application operates in the background without being visible.


## Using Project From Another Application: Early Binding

Early binding has better performance because it loads the type library at design time. To use early binding, you must set a reference to the Project application from the application you are working in. For example, in the VBE for a Word document, click  **References** on the **Tools** menu, scroll through the **Available References** list, and then choose the **Microsoft Project 15.0 Object Library** checkbox.

The following example opens a project from another application such as Excel , adds a task, and then saves and closes the project. 




```
Sub ModifyProject_Early() 
    Dim pjApp As MSProject.Application 
    Set pjApp = New MSProject.Application 
    pjApp.Visible = True 
    pjApp.FileOpen "Clocks.mpp" 
    pjApp.ActiveProject.Tasks.Add "Wind clocks" 
    pjApp.FileSave 
    pjApp.FileClose 
    pjApp.Quit 
End Sub
```


## Remarks




 **Important**  For application-level events, register event handlers  _after_ you set `Application.Visible = True`.



If you instantiate Project from another application and register an application-level event before setting the  **Visible** property of the **Application** object to **True**, the properties and methods of child objects of **Application** do not work. For example, `Application.ActiveProject.Name` is not accessible.

Many of the properties and methods that return the most common user-interface objects, such as the active project—represented by the  **[ActiveProject](http://msdn.microsoft.com/library/application-activeproject-property-project%28Office.15%29.aspx)** property—can be used without the **Application** object qualifier. For example, instead of writing `Application.ActiveProject.Visible = True` you can write `ActiveProject.Visible = True`


## Events



|**Name**|
|:-----|
|[AfterCubeBuilt](http://msdn.microsoft.com/library/application-aftercubebuilt-event-project%28Office.15%29.aspx)|
|[ApplicationBeforeClose](http://msdn.microsoft.com/library/application-applicationbeforeclose-event-project%28Office.15%29.aspx)|
|[ConnectionStatusChanged](http://msdn.microsoft.com/library/application-connectionstatuschanged-event-project%28Office.15%29.aspx)|
|[IsFunctionalitySupported](http://msdn.microsoft.com/library/application-isfunctionalitysupported-event-project%28Office.15%29.aspx)|
|[JobCompleted](http://msdn.microsoft.com/library/application-jobcompleted-event-project%28Office.15%29.aspx)|
|[JobStart](http://msdn.microsoft.com/library/application-jobstart-event-project%28Office.15%29.aspx)|
|[LoadWebPage](http://msdn.microsoft.com/library/application-loadwebpage-event-project%28Office.15%29.aspx)|
|[LoadWebPane](http://msdn.microsoft.com/library/application-loadwebpane-event-project%28Office.15%29.aspx)|
|[NewProject](http://msdn.microsoft.com/library/application-newproject-event-project%28Office.15%29.aspx)|
|[OnUndoOrRedo](http://msdn.microsoft.com/library/application-onundoorredo-event-project%28Office.15%29.aspx)|
|[PaneActivate](http://msdn.microsoft.com/library/application-paneactivate-event-project%28Office.15%29.aspx)|
|[ProjectAfterSave](http://msdn.microsoft.com/library/application-projectaftersave-event-project%28Office.15%29.aspx)|
|[ProjectAssignmentNew](http://msdn.microsoft.com/library/application-projectassignmentnew-event-project%28Office.15%29.aspx)|
|[ProjectBeforeAssignmentChange](http://msdn.microsoft.com/library/application-projectbeforeassignmentchange-event-project%28Office.15%29.aspx)|
|[ProjectBeforeAssignmentChange2](http://msdn.microsoft.com/library/application-projectbeforeassignmentchange2-event-project%28Office.15%29.aspx)|
|[ProjectBeforeAssignmentDelete](http://msdn.microsoft.com/library/application-projectbeforeassignmentdelete-event-project%28Office.15%29.aspx)|
|[ProjectBeforeAssignmentDelete2](http://msdn.microsoft.com/library/application-projectbeforeassignmentdelete2-event-project%28Office.15%29.aspx)|
|[ProjectBeforeAssignmentNew](http://msdn.microsoft.com/library/application-projectbeforeassignmentnew-event-project%28Office.15%29.aspx)|
|[ProjectBeforeAssignmentNew2](http://msdn.microsoft.com/library/application-projectbeforeassignmentnew2-event-project%28Office.15%29.aspx)|
|[ProjectBeforeClearBaseline](http://msdn.microsoft.com/library/application-projectbeforeclearbaseline-event-project%28Office.15%29.aspx)|
|[ProjectBeforeClose](http://msdn.microsoft.com/library/application-projectbeforeclose-event-project%28Office.15%29.aspx)|
|[ProjectBeforeClose2](http://msdn.microsoft.com/library/application-projectbeforeclose2-event-project%28Office.15%29.aspx)|
|[ProjectBeforePrint](http://msdn.microsoft.com/library/application-projectbeforeprint-event-project%28Office.15%29.aspx)|
|[ProjectBeforePrint2](http://msdn.microsoft.com/library/application-projectbeforeprint2-event-project%28Office.15%29.aspx)|
|[ProjectBeforePublish](http://msdn.microsoft.com/library/application-projectbeforepublish-event-project%28Office.15%29.aspx)|
|[ProjectBeforeResourceChange](http://msdn.microsoft.com/library/application-projectbeforeresourcechange-event-project%28Office.15%29.aspx)|
|[ProjectBeforeResourceChange2](http://msdn.microsoft.com/library/application-projectbeforeresourcechange2-event-project%28Office.15%29.aspx)|
|[ProjectBeforeResourceDelete](http://msdn.microsoft.com/library/application-projectbeforeresourcedelete-event-project%28Office.15%29.aspx)|
|[ProjectBeforeResourceDelete2](http://msdn.microsoft.com/library/application-projectbeforeresourcedelete2-event-project%28Office.15%29.aspx)|
|[ProjectBeforeResourceNew](http://msdn.microsoft.com/library/application-projectbeforeresourcenew-event-project%28Office.15%29.aspx)|
|[ProjectBeforeResourceNew2](http://msdn.microsoft.com/library/application-projectbeforeresourcenew2-event-project%28Office.15%29.aspx)|
|[ProjectBeforeSave](http://msdn.microsoft.com/library/application-projectbeforesave-event-project%28Office.15%29.aspx)|
|[ProjectBeforeSave2](http://msdn.microsoft.com/library/application-projectbeforesave2-event-project%28Office.15%29.aspx)|
|[ProjectBeforeSaveBaseline](http://msdn.microsoft.com/library/application-projectbeforesavebaseline-event-project%28Office.15%29.aspx)|
|[ProjectBeforeTaskChange](http://msdn.microsoft.com/library/application-projectbeforetaskchange-event-project%28Office.15%29.aspx)|
|[ProjectBeforeTaskChange2](http://msdn.microsoft.com/library/application-projectbeforetaskchange2-event-project%28Office.15%29.aspx)|
|[ProjectBeforeTaskDelete](http://msdn.microsoft.com/library/application-projectbeforetaskdelete-event-project%28Office.15%29.aspx)|
|[ProjectBeforeTaskDelete2](http://msdn.microsoft.com/library/application-projectbeforetaskdelete2-event-project%28Office.15%29.aspx)|
|[ProjectBeforeTaskNew](http://msdn.microsoft.com/library/application-projectbeforetasknew-event-project%28Office.15%29.aspx)|
|[ProjectBeforeTaskNew2](http://msdn.microsoft.com/library/application-projectbeforetasknew2-event-project%28Office.15%29.aspx)|
|[ProjectCalculate](http://msdn.microsoft.com/library/application-projectcalculate-event-project%28Office.15%29.aspx)|
|[ProjectResourceNew](http://msdn.microsoft.com/library/application-projectresourcenew-event-project%28Office.15%29.aspx)|
|[ProjectTaskNew](http://msdn.microsoft.com/library/application-projecttasknew-event-project%28Office.15%29.aspx)|
|[SaveCompletedToServer](http://msdn.microsoft.com/library/application-savecompletedtoserver-event-project%28Office.15%29.aspx)|
|[SaveStartingToServer](http://msdn.microsoft.com/library/application-savestartingtoserver-event-project%28Office.15%29.aspx)|
|[SecondaryViewChange](http://msdn.microsoft.com/library/application-secondaryviewchange-event-project%28Office.15%29.aspx)|
|[WindowActivate](http://msdn.microsoft.com/library/application-windowactivate-event-project%28Office.15%29.aspx)|
|[WindowBeforeViewChange](http://msdn.microsoft.com/library/application-windowbeforeviewchange-event-project%28Office.15%29.aspx)|
|[WindowDeactivate](http://msdn.microsoft.com/library/application-windowdeactivate-event-project%28Office.15%29.aspx)|
|[WindowGoalAreaChange](http://msdn.microsoft.com/library/application-windowgoalareachange-event-project%28Office.15%29.aspx)|
|[WindowSelectionChange](http://msdn.microsoft.com/library/application-windowselectionchange-event-project%28Office.15%29.aspx)|
|[WindowSidepaneDisplayChange](http://msdn.microsoft.com/library/application-windowsidepanedisplaychange-event-project%28Office.15%29.aspx)|
|[WindowSidepaneTaskChange](http://msdn.microsoft.com/library/application-windowsidepanetaskchange-event-project%28Office.15%29.aspx)|
|[WindowViewChange](http://msdn.microsoft.com/library/application-windowviewchange-event-project%28Office.15%29.aspx)|
|[WorkpaneDisplayChange](http://msdn.microsoft.com/library/application-workpanedisplaychange-event-project%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[About](http://msdn.microsoft.com/library/application-about-method-project%28Office.15%29.aspx)|
|[ActivateMicrosoftApp](http://msdn.microsoft.com/library/application-activatemicrosoftapp-method-project%28Office.15%29.aspx)|
|[AddNewColumn](http://msdn.microsoft.com/library/application-addnewcolumn-method-project%28Office.15%29.aspx)|
|[AddProgressLine](http://msdn.microsoft.com/library/application-addprogressline-method-project%28Office.15%29.aspx)|
|[AddResourcesFromProjectServer](http://msdn.microsoft.com/library/application-addresourcesfromprojectserver-method-project%28Office.15%29.aspx)|
|[AddSiteColumn](http://msdn.microsoft.com/library/application-addsitecolumn-method-project%28Office.15%29.aspx)|
|[AfterUnloadWebBrowserControl](http://msdn.microsoft.com/library/application-afterunloadwebbrowsercontrol-method-project%28Office.15%29.aspx)|
|[Alerts](http://msdn.microsoft.com/library/application-alerts-method-project%28Office.15%29.aspx)|
|[AlignTableCellBottom](http://msdn.microsoft.com/library/application-aligntablecellbottom-method-project%28Office.15%29.aspx)|
|[AlignTableCellTop](http://msdn.microsoft.com/library/application-aligntablecelltop-method-project%28Office.15%29.aspx)|
|[AlignTableCellVerticalCenter](http://msdn.microsoft.com/library/application-aligntablecellverticalcenter-method-project%28Office.15%29.aspx)|
|[AppExecute](http://msdn.microsoft.com/library/application-appexecute-method-project%28Office.15%29.aspx)|
|[ApplyReport](http://msdn.microsoft.com/library/application-applyreport-method-project%28Office.15%29.aspx)|
|[ApplyReportLayoutTemplate](http://msdn.microsoft.com/library/application-applyreportlayouttemplate-method-project%28Office.15%29.aspx)|
|[AppMaximize](http://msdn.microsoft.com/library/application-appmaximize-method-project%28Office.15%29.aspx)|
|[AppMinimize](http://msdn.microsoft.com/library/application-appminimize-method-project%28Office.15%29.aspx)|
|[AppMove](http://msdn.microsoft.com/library/application-appmove-method-project%28Office.15%29.aspx)|
|[AppRestore](http://msdn.microsoft.com/library/application-apprestore-method-project%28Office.15%29.aspx)|
|[AppSize](http://msdn.microsoft.com/library/application-appsize-method-project%28Office.15%29.aspx)|
|[AutoCorrect](http://msdn.microsoft.com/library/application-autocorrect-method-project%28Office.15%29.aspx)|
|[AutoFilter](http://msdn.microsoft.com/library/application-autofilter-method-project%28Office.15%29.aspx)|
|[AutoSaveToGlobal](http://msdn.microsoft.com/library/application-autosavetoglobal-method-project%28Office.15%29.aspx)|
|[BarBoxFormat](http://msdn.microsoft.com/library/application-barboxformat-method-project%28Office.15%29.aspx)|
|[BarBoxStyles](http://msdn.microsoft.com/library/application-barboxstyles-method-project%28Office.15%29.aspx)|
|[BarRounding](http://msdn.microsoft.com/library/application-barrounding-method-project%28Office.15%29.aspx)|
|[BaseCalendarCreate](http://msdn.microsoft.com/library/application-basecalendarcreate-method-project%28Office.15%29.aspx)|
|[BaseCalendarDelete](http://msdn.microsoft.com/library/application-basecalendardelete-method-project%28Office.15%29.aspx)|
|[BaseCalendarEditDays](http://msdn.microsoft.com/library/application-basecalendareditdays-method-project%28Office.15%29.aspx)|
|[BaseCalendarRename](http://msdn.microsoft.com/library/application-basecalendarrename-method-project%28Office.15%29.aspx)|
|[BaseCalendarReset](http://msdn.microsoft.com/library/application-basecalendarreset-method-project%28Office.15%29.aspx)|
|[BaseCalendars](http://msdn.microsoft.com/library/application-basecalendars-method-project%28Office.15%29.aspx)|
|[BaselineClear](http://msdn.microsoft.com/library/application-baselineclear-method-project%28Office.15%29.aspx)|
|[BaselineSave](http://msdn.microsoft.com/library/application-baselinesave-method-project%28Office.15%29.aspx)|
|[BoxAlign](http://msdn.microsoft.com/library/application-boxalign-method-project%28Office.15%29.aspx)|
|[BoxCellEdit](http://msdn.microsoft.com/library/application-boxcelledit-method-project%28Office.15%29.aspx)|
|[BoxCellEditEx](http://msdn.microsoft.com/library/application-boxcelleditex-method-project%28Office.15%29.aspx)|
|[BoxCellLayout](http://msdn.microsoft.com/library/application-boxcelllayout-method-project%28Office.15%29.aspx)|
|[BoxDataTemplate](http://msdn.microsoft.com/library/application-boxdatatemplate-method-project%28Office.15%29.aspx)|
|[BoxFormat](http://msdn.microsoft.com/library/application-boxformat-method-project%28Office.15%29.aspx)|
|[BoxFormatEx](http://msdn.microsoft.com/library/application-boxformatex-method-project%28Office.15%29.aspx)|
|[BoxGetXPosition](http://msdn.microsoft.com/library/application-boxgetxposition-method-project%28Office.15%29.aspx)|
|[BoxGetYPosition](http://msdn.microsoft.com/library/application-boxgetyposition-method-project%28Office.15%29.aspx)|
|[BoxLayout](http://msdn.microsoft.com/library/application-boxlayout-method-project%28Office.15%29.aspx)|
|[BoxLayoutEx](http://msdn.microsoft.com/library/application-boxlayoutex-method-project%28Office.15%29.aspx)|
|[BoxLinkLabelsShow](http://msdn.microsoft.com/library/application-boxlinklabelsshow-method-project%28Office.15%29.aspx)|
|[BoxLinks](http://msdn.microsoft.com/library/application-boxlinks-method-project%28Office.15%29.aspx)|
|[BoxLinksEx](http://msdn.microsoft.com/library/application-boxlinksex-method-project%28Office.15%29.aspx)|
|[BoxLinkStyleToggle](http://msdn.microsoft.com/library/application-boxlinkstyletoggle-method-project%28Office.15%29.aspx)|
|[BoxProgressMarksShow](http://msdn.microsoft.com/library/application-boxprogressmarksshow-method-project%28Office.15%29.aspx)|
|[BoxSet](http://msdn.microsoft.com/library/application-boxset-method-project%28Office.15%29.aspx)|
|[BoxShowHideFields](http://msdn.microsoft.com/library/application-boxshowhidefields-method-project%28Office.15%29.aspx)|
|[BoxStylesEdit](http://msdn.microsoft.com/library/application-boxstylesedit-method-project%28Office.15%29.aspx)|
|[BoxStylesEditEx](http://msdn.microsoft.com/library/application-boxstyleseditex-method-project%28Office.15%29.aspx)|
|[BoxZoom](http://msdn.microsoft.com/library/application-boxzoom-method-project%28Office.15%29.aspx)|
|[CacheSettings](http://msdn.microsoft.com/library/application-cachesettings-method-project%28Office.15%29.aspx)|
|[CacheStatus](http://msdn.microsoft.com/library/application-cachestatus-method-project%28Office.15%29.aspx)|
|[CalculateAll](http://msdn.microsoft.com/library/application-calculateall-method-project%28Office.15%29.aspx)|
|[CalculateProject](http://msdn.microsoft.com/library/application-calculateproject-method-project%28Office.15%29.aspx)|
|[CalendarBarStyles](http://msdn.microsoft.com/library/application-calendarbarstyles-method-project%28Office.15%29.aspx)|
|[CalendarBarStylesEdit](http://msdn.microsoft.com/library/application-calendarbarstylesedit-method-project%28Office.15%29.aspx)|
|[CalendarBarStylesEditEx](http://msdn.microsoft.com/library/application-calendarbarstyleseditex-method-project%28Office.15%29.aspx)|
|[CalendarBestFitWeekHeight](http://msdn.microsoft.com/library/application-calendarbestfitweekheight-method-project%28Office.15%29.aspx)|
|[CalendarDateBoxes](http://msdn.microsoft.com/library/application-calendardateboxes-method-project%28Office.15%29.aspx)|
|[CalendarDateBoxesEx](http://msdn.microsoft.com/library/application-calendardateboxesex-method-project%28Office.15%29.aspx)|
|[CalendarDateShading](http://msdn.microsoft.com/library/application-calendardateshading-method-project%28Office.15%29.aspx)|
|[CalendarDateShadingEdit](http://msdn.microsoft.com/library/application-calendardateshadingedit-method-project%28Office.15%29.aspx)|
|[CalendarDateShadingEditEx](http://msdn.microsoft.com/library/application-calendardateshadingeditex-method-project%28Office.15%29.aspx)|
|[CalendarLayout](http://msdn.microsoft.com/library/application-calendarlayout-method-project%28Office.15%29.aspx)|
|[CalendarShowBarSplits](http://msdn.microsoft.com/library/application-calendarshowbarsplits-method-project%28Office.15%29.aspx)|
|[CalendarTaskList](http://msdn.microsoft.com/library/application-calendartasklist-method-project%28Office.15%29.aspx)|
|[CalendarTimescale](http://msdn.microsoft.com/library/application-calendartimescale-method-project%28Office.15%29.aspx)|
|[CalendarWeekHeadingsEx](http://msdn.microsoft.com/library/application-calendarweekheadingsex-method-project%28Office.15%29.aspx)|
|[ChangeColumnDataType](http://msdn.microsoft.com/library/application-changecolumndatatype-method-project%28Office.15%29.aspx)|
|[ChangeStatusDate](http://msdn.microsoft.com/library/application-changestatusdate-method-project%28Office.15%29.aspx)|
|[ChangeWorkingTimeEx](http://msdn.microsoft.com/library/application-changeworkingtimeex-method-project%28Office.15%29.aspx)|
|[CheckField](http://msdn.microsoft.com/library/application-checkfield-method-project%28Office.15%29.aspx)|
|[CheckIn](http://msdn.microsoft.com/library/application-checkin-method-project%28Office.15%29.aspx)|
|[CheckOut](http://msdn.microsoft.com/library/application-checkout-method-project%28Office.15%29.aspx)|
|[CheckResourceErrors](http://msdn.microsoft.com/library/application-checkresourceerrors-method-project%28Office.15%29.aspx)|
|[CheckTaskErrors](http://msdn.microsoft.com/library/application-checktaskerrors-method-project%28Office.15%29.aspx)|
|[CleanupCache](http://msdn.microsoft.com/library/application-cleanupcache-method-project%28Office.15%29.aspx)|
|[CleanupProjectFromCache](http://msdn.microsoft.com/library/application-cleanupprojectfromcache-method-project%28Office.15%29.aspx)|
|[ClearConstraint](http://msdn.microsoft.com/library/application-clearconstraint-method-project%28Office.15%29.aspx)|
|[CloseComparison](http://msdn.microsoft.com/library/application-closecomparison-method-project%28Office.15%29.aspx)|
|[CloseUndoTransaction](http://msdn.microsoft.com/library/application-closeundotransaction-method-project%28Office.15%29.aspx)|
|[ColumnAlignment](http://msdn.microsoft.com/library/application-columnalignment-method-project%28Office.15%29.aspx)|
|[ColumnBestFit](http://msdn.microsoft.com/library/application-columnbestfit-method-project%28Office.15%29.aspx)|
|[ColumnDelete](http://msdn.microsoft.com/library/application-columndelete-method-project%28Office.15%29.aspx)|
|[ColumnEdit](http://msdn.microsoft.com/library/application-columnedit-method-project%28Office.15%29.aspx)|
|[ColumnInsert](http://msdn.microsoft.com/library/application-columninsert-method-project%28Office.15%29.aspx)|
|[ComAddInsDialog](http://msdn.microsoft.com/library/application-comaddinsdialog-method-project%28Office.15%29.aspx)|
|[CommitmentsPane](http://msdn.microsoft.com/library/application-commitmentspane-method-project%28Office.15%29.aspx)|
|[CompareProjectsLegendToggle](http://msdn.microsoft.com/library/application-compareprojectslegendtoggle-method-project%28Office.15%29.aspx)|
|[CompareProjectVersions](http://msdn.microsoft.com/library/application-compareprojectversions-method-project%28Office.15%29.aspx)|
|[ConsolidateProjects](http://msdn.microsoft.com/library/application-consolidateprojects-method-project%28Office.15%29.aspx)|
|[ConvertHangulToHanja](http://msdn.microsoft.com/library/application-converthangultohanja-method-project%28Office.15%29.aspx)|
|[CopyReport](http://msdn.microsoft.com/library/application-copyreport-method-project%28Office.15%29.aspx)|
|[CreateComparisonReport](http://msdn.microsoft.com/library/application-createcomparisonreport-method-project%28Office.15%29.aspx)|
|[CreateEnterpriseCalendar](http://msdn.microsoft.com/library/application-createenterprisecalendar-method-project%28Office.15%29.aspx)|
|[CreateProjectSite](http://msdn.microsoft.com/library/application-createprojectsite-method-project%28Office.15%29.aspx)|
|[CustomFieldDelete](http://msdn.microsoft.com/library/application-customfielddelete-method-project%28Office.15%29.aspx)|
|[CustomFieldGetFormula](http://msdn.microsoft.com/library/application-customfieldgetformula-method-project%28Office.15%29.aspx)|
|[CustomFieldGetName](http://msdn.microsoft.com/library/application-customfieldgetname-method-project%28Office.15%29.aspx)|
|[CustomFieldIndicatorAdd](http://msdn.microsoft.com/library/application-customfieldindicatoradd-method-project%28Office.15%29.aspx)|
|[CustomFieldIndicatorDelete](http://msdn.microsoft.com/library/application-customfieldindicatordelete-method-project%28Office.15%29.aspx)|
|[CustomFieldIndicators](http://msdn.microsoft.com/library/application-customfieldindicators-method-project%28Office.15%29.aspx)|
|[CustomFieldMappingDialog](http://msdn.microsoft.com/library/application-customfieldmappingdialog-method-project%28Office.15%29.aspx)|
|[CustomFieldPropertiesEx](http://msdn.microsoft.com/library/application-customfieldpropertiesex-method-project%28Office.15%29.aspx)|
|[CustomFieldRename](http://msdn.microsoft.com/library/application-customfieldrename-method-project%28Office.15%29.aspx)|
|[CustomFieldSetFormula](http://msdn.microsoft.com/library/application-customfieldsetformula-method-project%28Office.15%29.aspx)|
|[CustomFieldValueList](http://msdn.microsoft.com/library/application-customfieldvaluelist-method-project%28Office.15%29.aspx)|
|[CustomFieldValueListAdd](http://msdn.microsoft.com/library/application-customfieldvaluelistadd-method-project%28Office.15%29.aspx)|
|[CustomFieldValueListDelete](http://msdn.microsoft.com/library/application-customfieldvaluelistdelete-method-project%28Office.15%29.aspx)|
|[CustomFieldValueListGetItem](http://msdn.microsoft.com/library/application-customfieldvaluelistgetitem-method-project%28Office.15%29.aspx)|
|[CustomForms](http://msdn.microsoft.com/library/application-customforms-method-project%28Office.15%29.aspx)|
|[CustomizeField](http://msdn.microsoft.com/library/application-customizefield-method-project%28Office.15%29.aspx)|
|[CustomizeIMEMode](http://msdn.microsoft.com/library/application-customizeimemode-method-project%28Office.15%29.aspx)|
|[CustomOutlineCodeEditEx](http://msdn.microsoft.com/library/application-customoutlinecodeeditex-method-project%28Office.15%29.aspx)|
|[DateAdd](http://msdn.microsoft.com/library/application-dateadd-method-project%28Office.15%29.aspx)|
|[DateDifference](http://msdn.microsoft.com/library/application-datedifference-method-project%28Office.15%29.aspx)|
|[DateFormat](http://msdn.microsoft.com/library/application-dateformat-method-project%28Office.15%29.aspx)|
|[DateSubtract](http://msdn.microsoft.com/library/application-datesubtract-method-project%28Office.15%29.aspx)|
|[DDEExecute](http://msdn.microsoft.com/library/application-ddeexecute-method-project%28Office.15%29.aspx)|
|[DDEInitiate](http://msdn.microsoft.com/library/application-ddeinitiate-method-project%28Office.15%29.aspx)|
|[DDELinksUpdate](http://msdn.microsoft.com/library/application-ddelinksupdate-method-project%28Office.15%29.aspx)|
|[DDEPasteLink](http://msdn.microsoft.com/library/application-ddepastelink-method-project%28Office.15%29.aspx)|
|[DDETerminate](http://msdn.microsoft.com/library/application-ddeterminate-method-project%28Office.15%29.aspx)|
|[DeleteFromDatabase](http://msdn.microsoft.com/library/application-deletefromdatabase-method-project%28Office.15%29.aspx)|
|[DependenciesPane](http://msdn.microsoft.com/library/application-dependenciespane-method-project%28Office.15%29.aspx)|
|[DetailsPaneToggle](http://msdn.microsoft.com/library/application-detailspanetoggle-method-project%28Office.15%29.aspx)|
|[DetailStylesAdd](http://msdn.microsoft.com/library/application-detailstylesadd-method-project%28Office.15%29.aspx)|
|[DetailStylesFormat](http://msdn.microsoft.com/library/application-detailstylesformat-method-project%28Office.15%29.aspx)|
|[DetailStylesFormatEx](http://msdn.microsoft.com/library/application-detailstylesformatex-method-project%28Office.15%29.aspx)|
|[DetailStylesProperties](http://msdn.microsoft.com/library/application-detailstylesproperties-method-project%28Office.15%29.aspx)|
|[DetailStylesRemove](http://msdn.microsoft.com/library/application-detailstylesremove-method-project%28Office.15%29.aspx)|
|[DetailStylesRemoveAll](http://msdn.microsoft.com/library/application-detailstylesremoveall-method-project%28Office.15%29.aspx)|
|[DetailStylesToggleItem](http://msdn.microsoft.com/library/application-detailstylestoggleitem-method-project%28Office.15%29.aspx)|
|[DisplaySharedWorkspace](http://msdn.microsoft.com/library/application-displaysharedworkspace-method-project%28Office.15%29.aspx)|
|[DistributeTableColumns](http://msdn.microsoft.com/library/application-distributetablecolumns-method-project%28Office.15%29.aspx)|
|[DistributeTableRows](http://msdn.microsoft.com/library/application-distributetablerows-method-project%28Office.15%29.aspx)|
|[DocClose](http://msdn.microsoft.com/library/application-docclose-method-project%28Office.15%29.aspx)|
|[DocMaximize](http://msdn.microsoft.com/library/application-docmaximize-method-project%28Office.15%29.aspx)|
|[DocMove](http://msdn.microsoft.com/library/application-docmove-method-project%28Office.15%29.aspx)|
|[DocRestore](http://msdn.microsoft.com/library/application-docrestore-method-project%28Office.15%29.aspx)|
|[DocSize](http://msdn.microsoft.com/library/application-docsize-method-project%28Office.15%29.aspx)|
|[DocumentExport](http://msdn.microsoft.com/library/application-documentexport-method-project%28Office.15%29.aspx)|
|[DocumentLibraryVersionsDialog](http://msdn.microsoft.com/library/application-documentlibraryversionsdialog-method-project%28Office.15%29.aspx)|
|[DrawingCreate](http://msdn.microsoft.com/library/application-drawingcreate-method-project%28Office.15%29.aspx)|
|[DrawingCycleColor](http://msdn.microsoft.com/library/application-drawingcyclecolor-method-project%28Office.15%29.aspx)|
|[DrawingMove](http://msdn.microsoft.com/library/application-drawingmove-method-project%28Office.15%29.aspx)|
|[DrawingProperties](http://msdn.microsoft.com/library/application-drawingproperties-method-project%28Office.15%29.aspx)|
|[DrawingReshape](http://msdn.microsoft.com/library/application-drawingreshape-method-project%28Office.15%29.aspx)|
|[DurationFormat](http://msdn.microsoft.com/library/application-durationformat-method-project%28Office.15%29.aspx)|
|[DurationValue](http://msdn.microsoft.com/library/application-durationvalue-method-project%28Office.15%29.aspx)|
|[EditClear](http://msdn.microsoft.com/library/application-editclear-method-project%28Office.15%29.aspx)|
|[EditClearFormats](http://msdn.microsoft.com/library/application-editclearformats-method-project%28Office.15%29.aspx)|
|[EditClearHyperlink](http://msdn.microsoft.com/library/application-editclearhyperlink-method-project%28Office.15%29.aspx)|
|[EditCopy](http://msdn.microsoft.com/library/application-editcopy-method-project%28Office.15%29.aspx)|
|[EditCopyPicture](http://msdn.microsoft.com/library/application-editcopypicture-method-project%28Office.15%29.aspx)|
|[EditCut](http://msdn.microsoft.com/library/application-editcut-method-project%28Office.15%29.aspx)|
|[EditDelete](http://msdn.microsoft.com/library/application-editdelete-method-project%28Office.15%29.aspx)|
|[EditEnterpriseCalendar](http://msdn.microsoft.com/library/application-editenterprisecalendar-method-project%28Office.15%29.aspx)|
|[EditGoTo](http://msdn.microsoft.com/library/application-editgoto-method-project%28Office.15%29.aspx)|
|[EditHyperlink](http://msdn.microsoft.com/library/application-edithyperlink-method-project%28Office.15%29.aspx)|
|[EditInsert](http://msdn.microsoft.com/library/application-editinsert-method-project%28Office.15%29.aspx)|
|[EditPaste](http://msdn.microsoft.com/library/application-editpaste-method-project%28Office.15%29.aspx)|
|[EditPasteAsHyperlink](http://msdn.microsoft.com/library/application-editpasteashyperlink-method-project%28Office.15%29.aspx)|
|[EditPasteSpecial](http://msdn.microsoft.com/library/application-editpastespecial-method-project%28Office.15%29.aspx)|
|[EditRedo](http://msdn.microsoft.com/library/application-editredo-method-project%28Office.15%29.aspx)|
|[EditTPStyle](http://msdn.microsoft.com/library/application-edittpstyle-method-project%28Office.15%29.aspx)|
|[EditUndo](http://msdn.microsoft.com/library/application-editundo-method-project%28Office.15%29.aspx)|
|[EnterpriseGlobalCheckOut](http://msdn.microsoft.com/library/application-enterpriseglobalcheckout-method-project%28Office.15%29.aspx)|
|[EnterpriseMakeServerURLTrusted](http://msdn.microsoft.com/library/application-enterprisemakeserverurltrusted-method-project%28Office.15%29.aspx)|
|[EnterpriseProjectDelete](http://msdn.microsoft.com/library/application-enterpriseprojectdelete-method-project%28Office.15%29.aspx)|
|[EnterpriseProjectImportWizard](http://msdn.microsoft.com/library/application-enterpriseprojectimportwizard-method-project%28Office.15%29.aspx)|
|[EnterpriseProjectProfiles](http://msdn.microsoft.com/library/application-enterpriseprojectprofiles-method-project%28Office.15%29.aspx)|
|[EnterpriseResourceGet](http://msdn.microsoft.com/library/application-enterpriseresourceget-method-project%28Office.15%29.aspx)|
|[EnterpriseResourcesImportEx](http://msdn.microsoft.com/library/application-enterpriseresourcesimportex-method-project%28Office.15%29.aspx)|
|[EnterpriseResourcesOpen](http://msdn.microsoft.com/library/application-enterpriseresourcesopen-method-project%28Office.15%29.aspx)|
|[EnterpriseResSubstitutionWizard](http://msdn.microsoft.com/library/application-enterpriseressubstitutionwizard-method-project%28Office.15%29.aspx)|
|[EnterpriseTeamBuilder](http://msdn.microsoft.com/library/application-enterpriseteambuilder-method-project%28Office.15%29.aspx)|
|[FieldConstantToFieldName](http://msdn.microsoft.com/library/application-fieldconstanttofieldname-method-project%28Office.15%29.aspx)|
|[FieldNameToFieldConstant](http://msdn.microsoft.com/library/application-fieldnametofieldconstant-method-project%28Office.15%29.aspx)|
|[FileCloseAllEx](http://msdn.microsoft.com/library/application-filecloseallex-method-project%28Office.15%29.aspx)|
|[FileCloseEx](http://msdn.microsoft.com/library/application-filecloseex-method-project%28Office.15%29.aspx)|
|[FileExit](http://msdn.microsoft.com/library/application-fileexit-method-project%28Office.15%29.aspx)|
|[FileLoadLast](http://msdn.microsoft.com/library/application-fileloadlast-method-project%28Office.15%29.aspx)|
|[FileNew](http://msdn.microsoft.com/library/application-filenew-method-project%28Office.15%29.aspx)|
|[FileOpenEx](http://msdn.microsoft.com/library/application-fileopenex-method-project%28Office.15%29.aspx)|
|[FileOpenOrCreate](http://msdn.microsoft.com/library/application-fileopenorcreate-method-project%28Office.15%29.aspx)|
|[FileOpenUsingBackstage](http://msdn.microsoft.com/library/application-fileopenusingbackstage-method-project%28Office.15%29.aspx)|
|[FilePageSetup](http://msdn.microsoft.com/library/application-filepagesetup-method-project%28Office.15%29.aspx)|
|[FilePageSetupCalendar](http://msdn.microsoft.com/library/application-filepagesetupcalendar-method-project%28Office.15%29.aspx)|
|[FilePageSetupCalendarText](http://msdn.microsoft.com/library/application-filepagesetupcalendartext-method-project%28Office.15%29.aspx)|
|[FilePageSetupCalendarTextEx](http://msdn.microsoft.com/library/application-filepagesetupcalendartextex-method-project%28Office.15%29.aspx)|
|[FilePageSetupFooter](http://msdn.microsoft.com/library/application-filepagesetupfooter-method-project%28Office.15%29.aspx)|
|[FilePageSetupHeader](http://msdn.microsoft.com/library/application-filepagesetupheader-method-project%28Office.15%29.aspx)|
|[FilePageSetupLegend](http://msdn.microsoft.com/library/application-filepagesetuplegend-method-project%28Office.15%29.aspx)|
|[FilePageSetupLegendEx](http://msdn.microsoft.com/library/application-filepagesetuplegendex-method-project%28Office.15%29.aspx)|
|[FilePageSetupMargins](http://msdn.microsoft.com/library/application-filepagesetupmargins-method-project%28Office.15%29.aspx)|
|[FilePageSetupPage](http://msdn.microsoft.com/library/application-filepagesetuppage-method-project%28Office.15%29.aspx)|
|[FilePageSetupView](http://msdn.microsoft.com/library/application-filepagesetupview-method-project%28Office.15%29.aspx)|
|[FilePrint](http://msdn.microsoft.com/library/application-fileprint-method-project%28Office.15%29.aspx)|
|[FilePrintPreview](http://msdn.microsoft.com/library/application-fileprintpreview-method-project%28Office.15%29.aspx)|
|[FilePrintSetup](http://msdn.microsoft.com/library/application-fileprintsetup-method-project%28Office.15%29.aspx)|
|[FileProperties](http://msdn.microsoft.com/library/application-fileproperties-method-project%28Office.15%29.aspx)|
|[FileSave](http://msdn.microsoft.com/library/application-filesave-method-project%28Office.15%29.aspx)|
|[FileSaveAs](http://msdn.microsoft.com/library/application-filesaveas-method-project%28Office.15%29.aspx)|
|[FileSaveOffline](http://msdn.microsoft.com/library/application-filesaveoffline-method-project%28Office.15%29.aspx)|
|[FileSaveWorkspace](http://msdn.microsoft.com/library/application-filesaveworkspace-method-project%28Office.15%29.aspx)|
|[FillAcross](http://msdn.microsoft.com/library/application-fillacross-method-project%28Office.15%29.aspx)|
|[FillDown](http://msdn.microsoft.com/library/application-filldown-method-project%28Office.15%29.aspx)|
|[FilterApply](http://msdn.microsoft.com/library/application-filterapply-method-project%28Office.15%29.aspx)|
|[FilterClear](http://msdn.microsoft.com/library/application-filterclear-method-project%28Office.15%29.aspx)|
|[FilterEdit](http://msdn.microsoft.com/library/application-filteredit-method-project%28Office.15%29.aspx)|
|[FilterNew](http://msdn.microsoft.com/library/application-filternew-method-project%28Office.15%29.aspx)|
|[Filters](http://msdn.microsoft.com/library/application-filters-method-project%28Office.15%29.aspx)|
|[FilterShowSummaryRows](http://msdn.microsoft.com/library/application-filtershowsummaryrows-method-project%28Office.15%29.aspx)|
|[Find](http://msdn.microsoft.com/library/application-find-method-project%28Office.15%29.aspx)|
|[FindEx](http://msdn.microsoft.com/library/application-findex-method-project%28Office.15%29.aspx)|
|[FindFile](http://msdn.microsoft.com/library/application-findfile-method-project%28Office.15%29.aspx)|
|[FindNext](http://msdn.microsoft.com/library/application-findnext-method-project%28Office.15%29.aspx)|
|[FindPrevious](http://msdn.microsoft.com/library/application-findprevious-method-project%28Office.15%29.aspx)|
|[FollowHyperlink](http://msdn.microsoft.com/library/application-followhyperlink-method-project%28Office.15%29.aspx)|
|[Font32Ex](http://msdn.microsoft.com/library/application-font32ex-method-project%28Office.15%29.aspx)|
|[FontBold](http://msdn.microsoft.com/library/application-fontbold-method-project%28Office.15%29.aspx)|
|[FontEx](http://msdn.microsoft.com/library/application-fontex-method-project%28Office.15%29.aspx)|
|[FontItalic](http://msdn.microsoft.com/library/application-fontitalic-method-project%28Office.15%29.aspx)|
|[FontStrikethrough](http://msdn.microsoft.com/library/application-fontstrikethrough-method-project%28Office.15%29.aspx)|
|[FontUnderLine](http://msdn.microsoft.com/library/application-fontunderline-method-project%28Office.15%29.aspx)|
|[Form](http://msdn.microsoft.com/library/application-form-method-project%28Office.15%29.aspx)|
|[FormatCopy](http://msdn.microsoft.com/library/application-formatcopy-method-project%28Office.15%29.aspx)|
|[FormatPainter](http://msdn.microsoft.com/library/application-formatpainter-method-project%28Office.15%29.aspx)|
|[FormatPaste](http://msdn.microsoft.com/library/application-formatpaste-method-project%28Office.15%29.aspx)|
|[FormViewShow](http://msdn.microsoft.com/library/application-formviewshow-method-project%28Office.15%29.aspx)|
|[GanttBarEditEx](http://msdn.microsoft.com/library/application-ganttbareditex-method-project%28Office.15%29.aspx)|
|[GanttBarFormat](http://msdn.microsoft.com/library/application-ganttbarformat-method-project%28Office.15%29.aspx)|
|[GanttBarFormatEx](http://msdn.microsoft.com/library/application-ganttbarformatex-method-project%28Office.15%29.aspx)|
|[GanttBarLinks](http://msdn.microsoft.com/library/application-ganttbarlinks-method-project%28Office.15%29.aspx)|
|[GanttBarSize](http://msdn.microsoft.com/library/application-ganttbarsize-method-project%28Office.15%29.aspx)|
|[GanttBarStyleBaseline](http://msdn.microsoft.com/library/application-ganttbarstylebaseline-method-project%28Office.15%29.aspx)|
|[GanttBarStyleCritical](http://msdn.microsoft.com/library/application-ganttbarstylecritical-method-project%28Office.15%29.aspx)|
|[GanttBarStyleDelete](http://msdn.microsoft.com/library/application-ganttbarstyledelete-method-project%28Office.15%29.aspx)|
|[GanttBarStyleEdit](http://msdn.microsoft.com/library/application-ganttbarstyleedit-method-project%28Office.15%29.aspx)|
|[GanttBarStyleLate](http://msdn.microsoft.com/library/application-ganttbarstylelate-method-project%28Office.15%29.aspx)|
|[GanttBarStyleSlack](http://msdn.microsoft.com/library/application-ganttbarstyleslack-method-project%28Office.15%29.aspx)|
|[GanttBarStyleSlippage](http://msdn.microsoft.com/library/application-ganttbarstyleslippage-method-project%28Office.15%29.aspx)|
|[GanttBarTextDateFormat](http://msdn.microsoft.com/library/application-ganttbartextdateformat-method-project%28Office.15%29.aspx)|
|[GanttChartWizard](http://msdn.microsoft.com/library/application-ganttchartwizard-method-project%28Office.15%29.aspx)|
|[GanttRollup](http://msdn.microsoft.com/library/application-ganttrollup-method-project%28Office.15%29.aspx)|
|[GanttShowBarSplits](http://msdn.microsoft.com/library/application-ganttshowbarsplits-method-project%28Office.15%29.aspx)|
|[GanttShowDrawings](http://msdn.microsoft.com/library/application-ganttshowdrawings-method-project%28Office.15%29.aspx)|
|[GetCellInfo](http://msdn.microsoft.com/library/application-getcellinfo-method-project%28Office.15%29.aspx)|
|[GetCurrentTheme](http://msdn.microsoft.com/library/application-getcurrenttheme-method-project%28Office.15%29.aspx)|
|[GetProjectServerSettingsEx](http://msdn.microsoft.com/library/application-getprojectserversettingsex-method-project%28Office.15%29.aspx)|
|[GetProjectServerVersion](http://msdn.microsoft.com/library/application-getprojectserverversion-method-project%28Office.15%29.aspx)|
|[GetRedoListCount](http://msdn.microsoft.com/library/application-getredolistcount-method-project%28Office.15%29.aspx)|
|[GetRedoListItem](http://msdn.microsoft.com/library/application-getredolistitem-method-project%28Office.15%29.aspx)|
|[GetThemedColor](http://msdn.microsoft.com/library/application-getthemedcolor-method-project%28Office.15%29.aspx)|
|[GetUndoListCount](http://msdn.microsoft.com/library/application-getundolistcount-method-project%28Office.15%29.aspx)|
|[GetUndoListItem](http://msdn.microsoft.com/library/application-getundolistitem-method-project%28Office.15%29.aspx)|
|[GoalAreaChange](http://msdn.microsoft.com/library/application-goalareachange-method-project%28Office.15%29.aspx)|
|[GoalAreaHighlight](http://msdn.microsoft.com/library/application-goalareahighlight-method-project%28Office.15%29.aspx)|
|[GoalAreaTaskHighlight](http://msdn.microsoft.com/library/application-goalareataskhighlight-method-project%28Office.15%29.aspx)|
|[GoToItemInVersions](http://msdn.microsoft.com/library/application-gotoiteminversions-method-project%28Office.15%29.aspx)|
|[GotoNextOverAllocation](http://msdn.microsoft.com/library/application-gotonextoverallocation-method-project%28Office.15%29.aspx)|
|[GotoTaskDates](http://msdn.microsoft.com/library/application-gototaskdates-method-project%28Office.15%29.aspx)|
|[Gridlines](http://msdn.microsoft.com/library/application-gridlines-method-project%28Office.15%29.aspx)|
|[GridlinesEdit](http://msdn.microsoft.com/library/application-gridlinesedit-method-project%28Office.15%29.aspx)|
|[GridlinesEditEx](http://msdn.microsoft.com/library/application-gridlineseditex-method-project%28Office.15%29.aspx)|
|[GroupApply](http://msdn.microsoft.com/library/application-groupapply-method-project%28Office.15%29.aspx)|
|[GroupBy](http://msdn.microsoft.com/library/application-groupby-method-project%28Office.15%29.aspx)|
|[GroupClear](http://msdn.microsoft.com/library/application-groupclear-method-project%28Office.15%29.aspx)|
|[GroupMaintainHierarchy](http://msdn.microsoft.com/library/application-groupmaintainhierarchy-method-project%28Office.15%29.aspx)|
|[GroupNew](http://msdn.microsoft.com/library/application-groupnew-method-project%28Office.15%29.aspx)|
|[Groups](http://msdn.microsoft.com/library/application-groups-method-project%28Office.15%29.aspx)|
|[HelpAbout](http://msdn.microsoft.com/library/application-helpabout-method-project%28Office.15%29.aspx)|
|[HelpAnswerWizard](http://msdn.microsoft.com/library/application-helpanswerwizard-method-project%28Office.15%29.aspx)|
|[HelpContents](http://msdn.microsoft.com/library/application-helpcontents-method-project%28Office.15%29.aspx)|
|[HelpLaunch](http://msdn.microsoft.com/library/application-helplaunch-method-project%28Office.15%29.aspx)|
|[HelpTechnicalSupport](http://msdn.microsoft.com/library/application-helptechnicalsupport-method-project%28Office.15%29.aspx)|
|[HighlightDrivenSuccessors](http://msdn.microsoft.com/library/application-highlightdrivensuccessors-method-project%28Office.15%29.aspx)|
|[HighlightDrivingPredecessors](http://msdn.microsoft.com/library/application-highlightdrivingpredecessors-method-project%28Office.15%29.aspx)|
|[HighlightPredecessors](http://msdn.microsoft.com/library/application-highlightpredecessors-method-project%28Office.15%29.aspx)|
|[HighlightSuccessors](http://msdn.microsoft.com/library/application-highlightsuccessors-method-project%28Office.15%29.aspx)|
|[ImportCommitment](http://msdn.microsoft.com/library/application-importcommitment-method-project%28Office.15%29.aspx)|
|[ImportOutlookTasks](http://msdn.microsoft.com/library/application-importoutlooktasks-method-project%28Office.15%29.aspx)|
|[InactivateTaskToggle](http://msdn.microsoft.com/library/application-inactivatetasktoggle-method-project%28Office.15%29.aspx)|
|[InformationDialog](http://msdn.microsoft.com/library/application-informationdialog-method-project%28Office.15%29.aspx)|
|[InsertBlankRow](http://msdn.microsoft.com/library/application-insertblankrow-method-project%28Office.15%29.aspx)|
|[InsertHyperlink](http://msdn.microsoft.com/library/application-inserthyperlink-method-project%28Office.15%29.aspx)|
|[InsertManualTask](http://msdn.microsoft.com/library/application-insertmanualtask-method-project%28Office.15%29.aspx)|
|[InsertMilestoneTask](http://msdn.microsoft.com/library/application-insertmilestonetask-method-project%28Office.15%29.aspx)|
|[InsertNotes](http://msdn.microsoft.com/library/application-insertnotes-method-project%28Office.15%29.aspx)|
|[InsertResource](http://msdn.microsoft.com/library/application-insertresource-method-project%28Office.15%29.aspx)|
|[InsertScheduledTask](http://msdn.microsoft.com/library/application-insertscheduledtask-method-project%28Office.15%29.aspx)|
|[InsertSummaryTask](http://msdn.microsoft.com/library/application-insertsummarytask-method-project%28Office.15%29.aspx)|
|[InsertTask](http://msdn.microsoft.com/library/application-inserttask-method-project%28Office.15%29.aspx)|
|[IsCommandEnabled](http://msdn.microsoft.com/library/application-iscommandenabled-method-project%28Office.15%29.aspx)|
|[IsOfficeTaskPaneVisible](http://msdn.microsoft.com/library/application-isofficetaskpanevisible-method-project%28Office.15%29.aspx)|
|[IsOffline](http://msdn.microsoft.com/library/application-isoffline-method-project%28Office.15%29.aspx)|
|[IsReducedFunctionalityMode](http://msdn.microsoft.com/library/application-isreducedfunctionalitymode-method-project%28Office.15%29.aspx)|
|[IsUndoingOrRedoing](http://msdn.microsoft.com/library/application-isundoingorredoing-method-project%28Office.15%29.aspx)|
|[IsURLTrusted](http://msdn.microsoft.com/library/application-isurltrusted-method-project%28Office.15%29.aspx)|
|[Layout](http://msdn.microsoft.com/library/application-layout-method-project%28Office.15%29.aspx)|
|[LayoutNow](http://msdn.microsoft.com/library/application-layoutnow-method-project%28Office.15%29.aspx)|
|[LayoutRelatedNow](http://msdn.microsoft.com/library/application-layoutrelatednow-method-project%28Office.15%29.aspx)|
|[LayoutSelectionNow](http://msdn.microsoft.com/library/application-layoutselectionnow-method-project%28Office.15%29.aspx)|
|[LevelingClear](http://msdn.microsoft.com/library/application-levelingclear-method-project%28Office.15%29.aspx)|
|[LevelingOptions](http://msdn.microsoft.com/library/application-levelingoptions-method-project%28Office.15%29.aspx)|
|[LevelingOptionsEx](http://msdn.microsoft.com/library/application-levelingoptionsex-method-project%28Office.15%29.aspx)|
|[LevelNow](http://msdn.microsoft.com/library/application-levelnow-method-project%28Office.15%29.aspx)|
|[LevelSelected](http://msdn.microsoft.com/library/application-levelselected-method-project%28Office.15%29.aspx)|
|[LinksBetweenProjects](http://msdn.microsoft.com/library/application-linksbetweenprojects-method-project%28Office.15%29.aspx)|
|[LinkTasks](http://msdn.microsoft.com/library/application-linktasks-method-project%28Office.15%29.aspx)|
|[LinkTasksEdit](http://msdn.microsoft.com/library/application-linktasksedit-method-project%28Office.15%29.aspx)|
|[LinkToTaskList](http://msdn.microsoft.com/library/application-linktotasklist-method-project%28Office.15%29.aspx)|
|[LoadWebBrowserControlEx](http://msdn.microsoft.com/library/application-loadwebbrowsercontrolex-method-project%28Office.15%29.aspx)|
|[LoadWebPaneControl](http://msdn.microsoft.com/library/application-loadwebpanecontrol-method-project%28Office.15%29.aspx)|
|[LocaleID](http://msdn.microsoft.com/library/application-localeid-method-project%28Office.15%29.aspx)|
|[LookUpTableAddEx](http://msdn.microsoft.com/library/application-lookuptableaddex-method-project%28Office.15%29.aspx)|
|[Macro](http://msdn.microsoft.com/library/application-macro-method-project%28Office.15%29.aspx)|
|[MacroSecurity](http://msdn.microsoft.com/library/application-macrosecurity-method-project%28Office.15%29.aspx)|
|[MacroShowCode](http://msdn.microsoft.com/library/application-macroshowcode-method-project%28Office.15%29.aspx)|
|[MacroShowVba](http://msdn.microsoft.com/library/application-macroshowvba-method-project%28Office.15%29.aspx)|
|[MailLogoff](http://msdn.microsoft.com/library/application-maillogoff-method-project%28Office.15%29.aspx)|
|[MailLogon](http://msdn.microsoft.com/library/application-maillogon-method-project%28Office.15%29.aspx)|
|[MailPostDocument](http://msdn.microsoft.com/library/application-mailpostdocument-method-project%28Office.15%29.aspx)|
|[MailRoutingSlip](http://msdn.microsoft.com/library/application-mailroutingslip-method-project%28Office.15%29.aspx)|
|[MailSend](http://msdn.microsoft.com/library/application-mailsend-method-project%28Office.15%29.aspx)|
|[MailSession](http://msdn.microsoft.com/library/application-mailsession-method-project%28Office.15%29.aspx)|
|[MailSystem](http://msdn.microsoft.com/library/application-mailsystem-method-project%28Office.15%29.aspx)|
|[MakeFieldEnterprise](http://msdn.microsoft.com/library/application-makefieldenterprise-method-project%28Office.15%29.aspx)|
|[MakeLocalCalendarEnterprise](http://msdn.microsoft.com/library/application-makelocalcalendarenterprise-method-project%28Office.15%29.aspx)|
|[ManageSiteColumns](http://msdn.microsoft.com/library/application-managesitecolumns-method-project%28Office.15%29.aspx)|
|[MapEdit](http://msdn.microsoft.com/library/application-mapedit-method-project%28Office.15%29.aspx)|
|[Message](http://msdn.microsoft.com/library/application-message-method-project%28Office.15%29.aspx)|
|[NewTasksStartOn](http://msdn.microsoft.com/library/application-newtasksstarton-method-project%28Office.15%29.aspx)|
|[ObjectChangeIcon](http://msdn.microsoft.com/library/application-objectchangeicon-method-project%28Office.15%29.aspx)|
|[ObjectConvert](http://msdn.microsoft.com/library/application-objectconvert-method-project%28Office.15%29.aspx)|
|[ObjectInsert](http://msdn.microsoft.com/library/application-objectinsert-method-project%28Office.15%29.aspx)|
|[ObjectLinks](http://msdn.microsoft.com/library/application-objectlinks-method-project%28Office.15%29.aspx)|
|[ObjectVerb](http://msdn.microsoft.com/library/application-objectverb-method-project%28Office.15%29.aspx)|
|[OfficeOnTheWeb](http://msdn.microsoft.com/library/application-officeontheweb-method-project%28Office.15%29.aspx)|
|[OfficeTaskPaneHide](http://msdn.microsoft.com/library/application-officetaskpanehide-method-project%28Office.15%29.aspx)|
|[OpenBrowser](http://msdn.microsoft.com/library/application-openbrowser-method-project%28Office.15%29.aspx)|
|[OpenFromSharePoint](http://msdn.microsoft.com/library/application-openfromsharepoint-method-project%28Office.15%29.aspx)|
|[OpenServerPage](http://msdn.microsoft.com/library/application-openserverpage-method-project%28Office.15%29.aspx)|
|[OpenUndoTransaction](http://msdn.microsoft.com/library/application-openundotransaction-method-project%28Office.15%29.aspx)|
|[OpenXML](http://msdn.microsoft.com/library/application-openxml-method-project%28Office.15%29.aspx)|
|[OptionsCalculation](http://msdn.microsoft.com/library/application-optionscalculation-method-project%28Office.15%29.aspx)|
|[OptionsCalendar](http://msdn.microsoft.com/library/application-optionscalendar-method-project%28Office.15%29.aspx)|
|[OptionsEditEx](http://msdn.microsoft.com/library/application-optionseditex-method-project%28Office.15%29.aspx)|
|[OptionsGeneralEx](http://msdn.microsoft.com/library/application-optionsgeneralex-method-project%28Office.15%29.aspx)|
|[OptionsInterfaceEx](http://msdn.microsoft.com/library/application-optionsinterfaceex-method-project%28Office.15%29.aspx)|
|[OptionsSave](http://msdn.microsoft.com/library/application-optionssave-method-project%28Office.15%29.aspx)|
|[OptionsSchedule](http://msdn.microsoft.com/library/application-optionsschedule-method-project%28Office.15%29.aspx)|
|[OptionsSecurityEx](http://msdn.microsoft.com/library/application-optionssecurityex-method-project%28Office.15%29.aspx)|
|[OptionsSecurityTab](http://msdn.microsoft.com/library/application-optionssecuritytab-method-project%28Office.15%29.aspx)|
|[OptionsSpelling](http://msdn.microsoft.com/library/application-optionsspelling-method-project%28Office.15%29.aspx)|
|[OptionsViewEx](http://msdn.microsoft.com/library/application-optionsviewex-method-project%28Office.15%29.aspx)|
|[Organizer](http://msdn.microsoft.com/library/application-organizer-method-project%28Office.15%29.aspx)|
|[OrganizerDeleteItem](http://msdn.microsoft.com/library/application-organizerdeleteitem-method-project%28Office.15%29.aspx)|
|[OrganizerMoveItem](http://msdn.microsoft.com/library/application-organizermoveitem-method-project%28Office.15%29.aspx)|
|[OrganizerRenameItem](http://msdn.microsoft.com/library/application-organizerrenameitem-method-project%28Office.15%29.aspx)|
|[OutlineHideSubTasks](http://msdn.microsoft.com/library/application-outlinehidesubtasks-method-project%28Office.15%29.aspx)|
|[OutlineIndent](http://msdn.microsoft.com/library/application-outlineindent-method-project%28Office.15%29.aspx)|
|[OutlineOutdent](http://msdn.microsoft.com/library/application-outlineoutdent-method-project%28Office.15%29.aspx)|
|[OutlineShowAllTasks](http://msdn.microsoft.com/library/application-outlineshowalltasks-method-project%28Office.15%29.aspx)|
|[OutlineShowSubTasks](http://msdn.microsoft.com/library/application-outlineshowsubtasks-method-project%28Office.15%29.aspx)|
|[OutlineShowTasks](http://msdn.microsoft.com/library/application-outlineshowtasks-method-project%28Office.15%29.aspx)|
|[OutlineSymbolsToggle](http://msdn.microsoft.com/library/application-outlinesymbolstoggle-method-project%28Office.15%29.aspx)|
|[PageBreakRemove](http://msdn.microsoft.com/library/application-pagebreakremove-method-project%28Office.15%29.aspx)|
|[PageBreakSet](http://msdn.microsoft.com/library/application-pagebreakset-method-project%28Office.15%29.aspx)|
|[PageBreaksRemoveAll](http://msdn.microsoft.com/library/application-pagebreaksremoveall-method-project%28Office.15%29.aspx)|
|[PageBreaksShow](http://msdn.microsoft.com/library/application-pagebreaksshow-method-project%28Office.15%29.aspx)|
|[PaneClose](http://msdn.microsoft.com/library/application-paneclose-method-project%28Office.15%29.aspx)|
|[PaneCreate](http://msdn.microsoft.com/library/application-panecreate-method-project%28Office.15%29.aspx)|
|[PaneNext](http://msdn.microsoft.com/library/application-panenext-method-project%28Office.15%29.aspx)|
|[PanZoomPanTo](http://msdn.microsoft.com/library/application-panzoompanto-method-project%28Office.15%29.aspx)|
|[PanZoomZoomTo](http://msdn.microsoft.com/library/application-panzoomzoomto-method-project%28Office.15%29.aspx)|
|[PasteAsPicture](http://msdn.microsoft.com/library/application-pasteaspicture-method-project%28Office.15%29.aspx)|
|[PasteDestFormatting](http://msdn.microsoft.com/library/application-pastedestformatting-method-project%28Office.15%29.aspx)|
|[PasteSourceFormatting](http://msdn.microsoft.com/library/application-pastesourceformatting-method-project%28Office.15%29.aspx)|
|[ProgressLines](http://msdn.microsoft.com/library/application-progresslines-method-project%28Office.15%29.aspx)|
|[ProjectCheckOut](http://msdn.microsoft.com/library/application-projectcheckout-method-project%28Office.15%29.aspx)|
|[ProjectMove](http://msdn.microsoft.com/library/application-projectmove-method-project%28Office.15%29.aspx)|
|[ProjectStatistics](http://msdn.microsoft.com/library/application-projectstatistics-method-project%28Office.15%29.aspx)|
|[ProjectSummaryInfo](http://msdn.microsoft.com/library/application-projectsummaryinfo-method-project%28Office.15%29.aspx)|
|[Publish](http://msdn.microsoft.com/library/application-publish-method-project%28Office.15%29.aspx)|
|[Quit](http://msdn.microsoft.com/library/application-quit-method-project%28Office.15%29.aspx)|
|[ReassignSelectedAssns](http://msdn.microsoft.com/library/application-reassignselectedassns-method-project%28Office.15%29.aspx)|
|[RecurringTaskInsert](http://msdn.microsoft.com/library/application-recurringtaskinsert-method-project%28Office.15%29.aspx)|
|[Redo](http://msdn.microsoft.com/library/application-redo-method-project%28Office.15%29.aspx)|
|[RegisterProject](http://msdn.microsoft.com/library/application-registerproject-method-project%28Office.15%29.aspx)|
|[ReminderSet](http://msdn.microsoft.com/library/application-reminderset-method-project%28Office.15%29.aspx)|
|[RemoveHighlight](http://msdn.microsoft.com/library/application-removehighlight-method-project%28Office.15%29.aspx)|
|[RenameReport](http://msdn.microsoft.com/library/application-renamereport-method-project%28Office.15%29.aspx)|
|[Replace](http://msdn.microsoft.com/library/application-replace-method-project%28Office.15%29.aspx)|
|[ReplaceEx](http://msdn.microsoft.com/library/application-replaceex-method-project%28Office.15%29.aspx)|
|[ReportPrint](http://msdn.microsoft.com/library/application-reportprint-method-project%28Office.15%29.aspx)|
|[ReportPrintPreview](http://msdn.microsoft.com/library/application-reportprintpreview-method-project%28Office.15%29.aspx)|
|[Reports](http://msdn.microsoft.com/library/application-reports-method-project%28Office.15%29.aspx)|
|[ReportsDialog](http://msdn.microsoft.com/library/application-reportsdialog-method-project%28Office.15%29.aspx)|
|[RequestProgressInformation](http://msdn.microsoft.com/library/application-requestprogressinformation-method-project%28Office.15%29.aspx)|
|[RescheduleToNextAvailable](http://msdn.microsoft.com/library/application-rescheduletonextavailable-method-project%28Office.15%29.aspx)|
|[ResetTPStyle](http://msdn.microsoft.com/library/application-resettpstyle-method-project%28Office.15%29.aspx)|
|[ResourceActiveDirectory](http://msdn.microsoft.com/library/application-resourceactivedirectory-method-project%28Office.15%29.aspx)|
|[ResourceAddressBook](http://msdn.microsoft.com/library/application-resourceaddressbook-method-project%28Office.15%29.aspx)|
|[ResourceAssignment](http://msdn.microsoft.com/library/application-resourceassignment-method-project%28Office.15%29.aspx)|
|[ResourceAssignmentDialog](http://msdn.microsoft.com/library/application-resourceassignmentdialog-method-project%28Office.15%29.aspx)|
|[ResourceCalendarEditDays](http://msdn.microsoft.com/library/application-resourcecalendareditdays-method-project%28Office.15%29.aspx)|
|[ResourceCalendarReset](http://msdn.microsoft.com/library/application-resourcecalendarreset-method-project%28Office.15%29.aspx)|
|[ResourceCalendars](http://msdn.microsoft.com/library/application-resourcecalendars-method-project%28Office.15%29.aspx)|
|[ResourceComparison](http://msdn.microsoft.com/library/application-resourcecomparison-method-project%28Office.15%29.aspx)|
|[ResourceDetails](http://msdn.microsoft.com/library/application-resourcedetails-method-project%28Office.15%29.aspx)|
|[ResourceGraphBarStyles](http://msdn.microsoft.com/library/application-resourcegraphbarstyles-method-project%28Office.15%29.aspx)|
|[ResourceGraphBarStylesEx](http://msdn.microsoft.com/library/application-resourcegraphbarstylesex-method-project%28Office.15%29.aspx)|
|[ResourceMappingDialog](http://msdn.microsoft.com/library/application-resourcemappingdialog-method-project%28Office.15%29.aspx)|
|[ResourceSharing](http://msdn.microsoft.com/library/application-resourcesharing-method-project%28Office.15%29.aspx)|
|[ResourceSharingPoolAction](http://msdn.microsoft.com/library/application-resourcesharingpoolaction-method-project%28Office.15%29.aspx)|
|[ResourceSharingPoolRefresh](http://msdn.microsoft.com/library/application-resourcesharingpoolrefresh-method-project%28Office.15%29.aspx)|
|[ResourceSharingPoolUpdate](http://msdn.microsoft.com/library/application-resourcesharingpoolupdate-method-project%28Office.15%29.aspx)|
|[ResourceWindowsAccount](http://msdn.microsoft.com/library/application-resourcewindowsaccount-method-project%28Office.15%29.aspx)|
|[RestoreSheetSelection](http://msdn.microsoft.com/library/application-restoresheetselection-method-project%28Office.15%29.aspx)|
|[RowClear](http://msdn.microsoft.com/library/application-rowclear-method-project%28Office.15%29.aspx)|
|[RowDelete](http://msdn.microsoft.com/library/application-rowdelete-method-project%28Office.15%29.aspx)|
|[RowInsert](http://msdn.microsoft.com/library/application-rowinsert-method-project%28Office.15%29.aspx)|
|[Run](http://msdn.microsoft.com/library/application-run-method-project%28Office.15%29.aspx)|
|[SaveForSharing](http://msdn.microsoft.com/library/application-saveforsharing-method-project%28Office.15%29.aspx)|
|[SaveSheetSelection](http://msdn.microsoft.com/library/application-savesheetselection-method-project%28Office.15%29.aspx)|
|[SegmentBorderColor](http://msdn.microsoft.com/library/application-segmentbordercolor-method-project%28Office.15%29.aspx)|
|[SegmentFillColor](http://msdn.microsoft.com/library/application-segmentfillcolor-method-project%28Office.15%29.aspx)|
|[SelectAll](http://msdn.microsoft.com/library/application-selectall-method-project%28Office.15%29.aspx)|
|[SelectBeginning](http://msdn.microsoft.com/library/application-selectbeginning-method-project%28Office.15%29.aspx)|
|[SelectCell](http://msdn.microsoft.com/library/application-selectcell-method-project%28Office.15%29.aspx)|
|[SelectCellDown](http://msdn.microsoft.com/library/application-selectcelldown-method-project%28Office.15%29.aspx)|
|[SelectCellLeft](http://msdn.microsoft.com/library/application-selectcellleft-method-project%28Office.15%29.aspx)|
|[SelectCellRight](http://msdn.microsoft.com/library/application-selectcellright-method-project%28Office.15%29.aspx)|
|[SelectCellUp](http://msdn.microsoft.com/library/application-selectcellup-method-project%28Office.15%29.aspx)|
|[SelectColumn](http://msdn.microsoft.com/library/application-selectcolumn-method-project%28Office.15%29.aspx)|
|[SelectEnd](http://msdn.microsoft.com/library/application-selectend-method-project%28Office.15%29.aspx)|
|[SelectionExtend](http://msdn.microsoft.com/library/application-selectionextend-method-project%28Office.15%29.aspx)|
|[SelectRange](http://msdn.microsoft.com/library/application-selectrange-method-project%28Office.15%29.aspx)|
|[SelectResourceCell](http://msdn.microsoft.com/library/application-selectresourcecell-method-project%28Office.15%29.aspx)|
|[SelectResourceColumn](http://msdn.microsoft.com/library/application-selectresourcecolumn-method-project%28Office.15%29.aspx)|
|[SelectResourceField](http://msdn.microsoft.com/library/application-selectresourcefield-method-project%28Office.15%29.aspx)|
|[SelectRow](http://msdn.microsoft.com/library/application-selectrow-method-project%28Office.15%29.aspx)|
|[SelectRowEnd](http://msdn.microsoft.com/library/application-selectrowend-method-project%28Office.15%29.aspx)|
|[SelectRowStart](http://msdn.microsoft.com/library/application-selectrowstart-method-project%28Office.15%29.aspx)|
|[SelectSheet](http://msdn.microsoft.com/library/application-selectsheet-method-project%28Office.15%29.aspx)|
|[SelectTable](http://msdn.microsoft.com/library/application-selecttable-method-project%28Office.15%29.aspx)|
|[SelectTaskAssns](http://msdn.microsoft.com/library/application-selecttaskassns-method-project%28Office.15%29.aspx)|
|[SelectTaskCell](http://msdn.microsoft.com/library/application-selecttaskcell-method-project%28Office.15%29.aspx)|
|[SelectTaskColumn](http://msdn.microsoft.com/library/application-selecttaskcolumn-method-project%28Office.15%29.aspx)|
|[SelectTaskField](http://msdn.microsoft.com/library/application-selecttaskfield-method-project%28Office.15%29.aspx)|
|[SelectTimescaleRange](http://msdn.microsoft.com/library/application-selecttimescalerange-method-project%28Office.15%29.aspx)|
|[SelectToEnd](http://msdn.microsoft.com/library/application-selecttoend-method-project%28Office.15%29.aspx)|
|[SelectTPLineHeight](http://msdn.microsoft.com/library/application-selecttplineheight-method-project%28Office.15%29.aspx)|
|[SelectTPTask](http://msdn.microsoft.com/library/application-selecttptask-method-project%28Office.15%29.aspx)|
|[ServiceOptionsDialog](http://msdn.microsoft.com/library/application-serviceoptionsdialog-method-project%28Office.15%29.aspx)|
|[SetActiveCell](http://msdn.microsoft.com/library/application-setactivecell-method-project%28Office.15%29.aspx)|
|[SetAutoFilter](http://msdn.microsoft.com/library/application-setautofilter-method-project%28Office.15%29.aspx)|
|[SetField](http://msdn.microsoft.com/library/application-setfield-method-project%28Office.15%29.aspx)|
|[SetLTRTable](http://msdn.microsoft.com/library/application-setltrtable-method-project%28Office.15%29.aspx)|
|[SetMatchingField](http://msdn.microsoft.com/library/application-setmatchingfield-method-project%28Office.15%29.aspx)|
|[SetResourceField](http://msdn.microsoft.com/library/application-setresourcefield-method-project%28Office.15%29.aspx)|
|[SetResourceFieldByID](http://msdn.microsoft.com/library/application-setresourcefieldbyid-method-project%28Office.15%29.aspx)|
|[SetRowHeight](http://msdn.microsoft.com/library/application-setrowheight-method-project%28Office.15%29.aspx)|
|[SetRTLTable](http://msdn.microsoft.com/library/application-setrtltable-method-project%28Office.15%29.aspx)|
|[SetShowTaskSuggestions](http://msdn.microsoft.com/library/application-setshowtasksuggestions-method-project%28Office.15%29.aspx)|
|[SetShowTaskWarnings](http://msdn.microsoft.com/library/application-setshowtaskwarnings-method-project%28Office.15%29.aspx)|
|[SetSidepaneStateButton](http://msdn.microsoft.com/library/application-setsidepanestatebutton-method-project%28Office.15%29.aspx)|
|[SetSplitBar](http://msdn.microsoft.com/library/application-setsplitbar-method-project%28Office.15%29.aspx)|
|[SetTaskField](http://msdn.microsoft.com/library/application-settaskfield-method-project%28Office.15%29.aspx)|
|[SetTaskFieldByID](http://msdn.microsoft.com/library/application-settaskfieldbyid-method-project%28Office.15%29.aspx)|
|[SetTaskMode](http://msdn.microsoft.com/library/application-settaskmode-method-project%28Office.15%29.aspx)|
|[SetTitleRowHeight](http://msdn.microsoft.com/library/application-settitlerowheight-method-project%28Office.15%29.aspx)|
|[SetTPField](http://msdn.microsoft.com/library/application-settpfield-method-project%28Office.15%29.aspx)|
|[ShareProjectOnline](http://msdn.microsoft.com/library/application-shareprojectonline-method-project%28Office.15%29.aspx)|
|[ShowAddNewColumn](http://msdn.microsoft.com/library/application-showaddnewcolumn-method-project%28Office.15%29.aspx)|
|[ShowIgnoredTaskWarnings](http://msdn.microsoft.com/library/application-showignoredtaskwarnings-method-project%28Office.15%29.aspx)|
|[ShowOSFTaskPane](http://msdn.microsoft.com/library/application-showosftaskpane-method-project%28Office.15%29.aspx)|
|[ShowReportDataPane](http://msdn.microsoft.com/library/application-showreportdatapane-method-project%28Office.15%29.aspx)|
|[SidepaneTaskChange](http://msdn.microsoft.com/library/application-sidepanetaskchange-method-project%28Office.15%29.aspx)|
|[SidepaneToggle](http://msdn.microsoft.com/library/application-sidepanetoggle-method-project%28Office.15%29.aspx)|
|[Sort](http://msdn.microsoft.com/library/application-sort-method-project%28Office.15%29.aspx)|
|[SpellCheckField](http://msdn.microsoft.com/library/application-spellcheckfield-method-project%28Office.15%29.aspx)|
|[SpellingCheck](http://msdn.microsoft.com/library/application-spellingcheck-method-project%28Office.15%29.aspx)|
|[SplitTask](http://msdn.microsoft.com/library/application-splittask-method-project%28Office.15%29.aspx)|
|[StopWebBrowserControlNavigation](http://msdn.microsoft.com/library/application-stopwebbrowsercontrolnavigation-method-project%28Office.15%29.aspx)|
|[SummaryResourceAssignmentsRefresh](http://msdn.microsoft.com/library/application-summaryresourceassignmentsrefresh-method-project%28Office.15%29.aspx)|
|[SummaryTasksShow](http://msdn.microsoft.com/library/application-summarytasksshow-method-project%28Office.15%29.aspx)|
|[SynchronizeWithSite](http://msdn.microsoft.com/library/application-synchronizewithsite-method-project%28Office.15%29.aspx)|
|[Table](http://msdn.microsoft.com/library/application-table-method-project%28Office.15%29.aspx)|
|[TableApply](http://msdn.microsoft.com/library/application-tableapply-method-project%28Office.15%29.aspx)|
|[TableCopy](http://msdn.microsoft.com/library/application-tablecopy-method-project%28Office.15%29.aspx)|
|[TableEdit](http://msdn.microsoft.com/library/application-tableedit-method-project%28Office.15%29.aspx)|
|[TableEditEx](http://msdn.microsoft.com/library/application-tableeditex-method-project%28Office.15%29.aspx)|
|[TableReset](http://msdn.microsoft.com/library/application-tablereset-method-project%28Office.15%29.aspx)|
|[Tables](http://msdn.microsoft.com/library/application-tables-method-project%28Office.15%29.aspx)|
|[TaskComparison](http://msdn.microsoft.com/library/application-taskcomparison-method-project%28Office.15%29.aspx)|
|[TaskDeliverableCreate](http://msdn.microsoft.com/library/application-taskdeliverablecreate-method-project%28Office.15%29.aspx)|
|[TaskDeliverableSync](http://msdn.microsoft.com/library/application-taskdeliverablesync-method-project%28Office.15%29.aspx)|
|[TaskDependencySync](http://msdn.microsoft.com/library/application-taskdependencysync-method-project%28Office.15%29.aspx)|
|[TaskDrivers](http://msdn.microsoft.com/library/application-taskdrivers-method-project%28Office.15%29.aspx)|
|[TaskInspector](http://msdn.microsoft.com/library/application-taskinspector-method-project%28Office.15%29.aspx)|
|[TaskMove](http://msdn.microsoft.com/library/application-taskmove-method-project%28Office.15%29.aspx)|
|[TaskMoveToStatusDate](http://msdn.microsoft.com/library/application-taskmovetostatusdate-method-project%28Office.15%29.aspx)|
|[TaskOnTimeline](http://msdn.microsoft.com/library/application-taskontimeline-method-project%28Office.15%29.aspx)|
|[TaskRespectLinks](http://msdn.microsoft.com/library/application-taskrespectlinks-method-project%28Office.15%29.aspx)|
|[TextStyles32Ex](http://msdn.microsoft.com/library/application-textstyles32ex-method-project%28Office.15%29.aspx)|
|[TextStylesEx](http://msdn.microsoft.com/library/application-textstylesex-method-project%28Office.15%29.aspx)|
|[TimelineExport](http://msdn.microsoft.com/library/application-timelineexport-method-project%28Office.15%29.aspx)|
|[TimelineFormat](http://msdn.microsoft.com/library/application-timelineformat-method-project%28Office.15%29.aspx)|
|[TimelineGotoSelectedTask](http://msdn.microsoft.com/library/application-timelinegotoselectedtask-method-project%28Office.15%29.aspx)|
|[TimelineInsertTask](http://msdn.microsoft.com/library/application-timelineinserttask-method-project%28Office.15%29.aspx)|
|[TimelineShowHide](http://msdn.microsoft.com/library/application-timelineshowhide-method-project%28Office.15%29.aspx)|
|[TimelineTextOnBar](http://msdn.microsoft.com/library/application-timelinetextonbar-method-project%28Office.15%29.aspx)|
|[TimelineViewToggle](http://msdn.microsoft.com/library/application-timelineviewtoggle-method-project%28Office.15%29.aspx)|
|[Timescale](http://msdn.microsoft.com/library/application-timescale-method-project%28Office.15%29.aspx)|
|[TimescaleEdit](http://msdn.microsoft.com/library/application-timescaleedit-method-project%28Office.15%29.aspx)|
|[TimescaleNonWorking](http://msdn.microsoft.com/library/application-timescalenonworking-method-project%28Office.15%29.aspx)|
|[TimescaleNonWorkingEx](http://msdn.microsoft.com/library/application-timescalenonworkingex-method-project%28Office.15%29.aspx)|
|[ToggleAssignments](http://msdn.microsoft.com/library/application-toggleassignments-method-project%28Office.15%29.aspx)|
|[ToggleChangeHighlighting](http://msdn.microsoft.com/library/application-togglechangehighlighting-method-project%28Office.15%29.aspx)|
|[TogglePreventResOveralloc](http://msdn.microsoft.com/library/application-togglepreventresoveralloc-method-project%28Office.15%29.aspx)|
|[ToggleResourceDetails](http://msdn.microsoft.com/library/application-toggleresourcedetails-method-project%28Office.15%29.aspx)|
|[ToggleTaskDetails](http://msdn.microsoft.com/library/application-toggletaskdetails-method-project%28Office.15%29.aspx)|
|[ToggleTPAutoExpand](http://msdn.microsoft.com/library/application-toggletpautoexpand-method-project%28Office.15%29.aspx)|
|[ToggleTPResourceExpand](http://msdn.microsoft.com/library/application-toggletpresourceexpand-method-project%28Office.15%29.aspx)|
|[ToggleTPUnassigned](http://msdn.microsoft.com/library/application-toggletpunassigned-method-project%28Office.15%29.aspx)|
|[ToggleTPUnscheduled](http://msdn.microsoft.com/library/application-toggletpunscheduled-method-project%28Office.15%29.aspx)|
|[Undo](http://msdn.microsoft.com/library/application-undo-method-project%28Office.15%29.aspx)|
|[UndoClear](http://msdn.microsoft.com/library/application-undoclear-method-project%28Office.15%29.aspx)|
|[UnlinkTasks](http://msdn.microsoft.com/library/application-unlinktasks-method-project%28Office.15%29.aspx)|
|[UnloadWebBrowserControl](http://msdn.microsoft.com/library/application-unloadwebbrowsercontrol-method-project%28Office.15%29.aspx)|
|[UpdateFromProjectServer](http://msdn.microsoft.com/library/application-updatefromprojectserver-method-project%28Office.15%29.aspx)|
|[UpdateProject](http://msdn.microsoft.com/library/application-updateproject-method-project%28Office.15%29.aspx)|
|[UpdateTasks](http://msdn.microsoft.com/library/application-updatetasks-method-project%28Office.15%29.aspx)|
|[UsageViewEntryEx](http://msdn.microsoft.com/library/application-usageviewentryex-method-project%28Office.15%29.aspx)|
|[ViewApply](http://msdn.microsoft.com/library/application-viewapply-method-project%28Office.15%29.aspx)|
|[ViewApplyEx](http://msdn.microsoft.com/library/application-viewapplyex-method-project%28Office.15%29.aspx)|
|[ViewBar](http://msdn.microsoft.com/library/application-viewbar-method-project%28Office.15%29.aspx)|
|[ViewCopy](http://msdn.microsoft.com/library/application-viewcopy-method-project%28Office.15%29.aspx)|
|[ViewEditCombination](http://msdn.microsoft.com/library/application-vieweditcombination-method-project%28Office.15%29.aspx)|
|[ViewEditSingle](http://msdn.microsoft.com/library/application-vieweditsingle-method-project%28Office.15%29.aspx)|
|[ViewReset](http://msdn.microsoft.com/library/application-viewreset-method-project%28Office.15%29.aspx)|
|[Views](http://msdn.microsoft.com/library/application-views-method-project%28Office.15%29.aspx)|
|[ViewsEx](http://msdn.microsoft.com/library/application-viewsex-method-project%28Office.15%29.aspx)|
|[ViewShowCost](http://msdn.microsoft.com/library/application-viewshowcost-method-project%28Office.15%29.aspx)|
|[ViewShowCumulativeCost](http://msdn.microsoft.com/library/application-viewshowcumulativecost-method-project%28Office.15%29.aspx)|
|[ViewShowCumulativeWork](http://msdn.microsoft.com/library/application-viewshowcumulativework-method-project%28Office.15%29.aspx)|
|[ViewShowNotes](http://msdn.microsoft.com/library/application-viewshownotes-method-project%28Office.15%29.aspx)|
|[ViewShowObjects](http://msdn.microsoft.com/library/application-viewshowobjects-method-project%28Office.15%29.aspx)|
|[ViewShowOverallocation](http://msdn.microsoft.com/library/application-viewshowoverallocation-method-project%28Office.15%29.aspx)|
|[ViewShowPeakUnits](http://msdn.microsoft.com/library/application-viewshowpeakunits-method-project%28Office.15%29.aspx)|
|[ViewShowPercentAllocation](http://msdn.microsoft.com/library/application-viewshowpercentallocation-method-project%28Office.15%29.aspx)|
|[ViewShowPredecessorsSuccessors](http://msdn.microsoft.com/library/application-viewshowpredecessorssuccessors-method-project%28Office.15%29.aspx)|
|[ViewShowRemainingAvailability](http://msdn.microsoft.com/library/application-viewshowremainingavailability-method-project%28Office.15%29.aspx)|
|[ViewShowResourcesPredecessors](http://msdn.microsoft.com/library/application-viewshowresourcespredecessors-method-project%28Office.15%29.aspx)|
|[ViewShowResourcesSuccessors](http://msdn.microsoft.com/library/application-viewshowresourcessuccessors-method-project%28Office.15%29.aspx)|
|[ViewShowSchedule](http://msdn.microsoft.com/library/application-viewshowschedule-method-project%28Office.15%29.aspx)|
|[ViewShowUnitAvailability](http://msdn.microsoft.com/library/application-viewshowunitavailability-method-project%28Office.15%29.aspx)|
|[ViewShowWork](http://msdn.microsoft.com/library/application-viewshowwork-method-project%28Office.15%29.aspx)|
|[ViewShowWorkAvailability](http://msdn.microsoft.com/library/application-viewshowworkavailability-method-project%28Office.15%29.aspx)|
|[VisualReports](http://msdn.microsoft.com/library/application-visualreports-method-project%28Office.15%29.aspx)|
|[VisualReportsEdit](http://msdn.microsoft.com/library/application-visualreportsedit-method-project%28Office.15%29.aspx)|
|[VisualReportsNewTemplate](http://msdn.microsoft.com/library/application-visualreportsnewtemplate-method-project%28Office.15%29.aspx)|
|[VisualReportsSaveCube](http://msdn.microsoft.com/library/application-visualreportssavecube-method-project%28Office.15%29.aspx)|
|[VisualReportsSaveDatabase](http://msdn.microsoft.com/library/application-visualreportssavedatabase-method-project%28Office.15%29.aspx)|
|[VisualReportsView](http://msdn.microsoft.com/library/application-visualreportsview-method-project%28Office.15%29.aspx)|
|[WBSCodeMaskEdit](http://msdn.microsoft.com/library/application-wbscodemaskedit-method-project%28Office.15%29.aspx)|
|[WBSCodeRenumber](http://msdn.microsoft.com/library/application-wbscoderenumber-method-project%28Office.15%29.aspx)|
|[WebAddToFavorites](http://msdn.microsoft.com/library/application-webaddtofavorites-method-project%28Office.15%29.aspx)|
|[WebCopyHyperlink](http://msdn.microsoft.com/library/application-webcopyhyperlink-method-project%28Office.15%29.aspx)|
|[WebGoBack](http://msdn.microsoft.com/library/application-webgoback-method-project%28Office.15%29.aspx)|
|[WebGoForward](http://msdn.microsoft.com/library/application-webgoforward-method-project%28Office.15%29.aspx)|
|[WebHideToolbars](http://msdn.microsoft.com/library/application-webhidetoolbars-method-project%28Office.15%29.aspx)|
|[WebOpenFavorites](http://msdn.microsoft.com/library/application-webopenfavorites-method-project%28Office.15%29.aspx)|
|[WebOpenHyperlink](http://msdn.microsoft.com/library/application-webopenhyperlink-method-project%28Office.15%29.aspx)|
|[WebOpenSearchPage](http://msdn.microsoft.com/library/application-webopensearchpage-method-project%28Office.15%29.aspx)|
|[WebOpenStartPage](http://msdn.microsoft.com/library/application-webopenstartpage-method-project%28Office.15%29.aspx)|
|[WebRefresh](http://msdn.microsoft.com/library/application-webrefresh-method-project%28Office.15%29.aspx)|
|[WebSetSearchPage](http://msdn.microsoft.com/library/application-websetsearchpage-method-project%28Office.15%29.aspx)|
|[WebSetStartPage](http://msdn.microsoft.com/library/application-websetstartpage-method-project%28Office.15%29.aspx)|
|[WebStopLoading](http://msdn.microsoft.com/library/application-webstoploading-method-project%28Office.15%29.aspx)|
|[WebToolbar](http://msdn.microsoft.com/library/application-webtoolbar-method-project%28Office.15%29.aspx)|
|[WindowActivate](http://msdn.microsoft.com/library/application-windowactivate-method-project%28Office.15%29.aspx)|
|[WindowArrangeAll](http://msdn.microsoft.com/library/application-windowarrangeall-method-project%28Office.15%29.aspx)|
|[WindowHide](http://msdn.microsoft.com/library/application-windowhide-method-project%28Office.15%29.aspx)|
|[WindowMoreWindows](http://msdn.microsoft.com/library/application-windowmorewindows-method-project%28Office.15%29.aspx)|
|[WindowNewWindow](http://msdn.microsoft.com/library/application-windownewwindow-method-project%28Office.15%29.aspx)|
|[WindowNext](http://msdn.microsoft.com/library/application-windownext-method-project%28Office.15%29.aspx)|
|[WindowPrev](http://msdn.microsoft.com/library/application-windowprev-method-project%28Office.15%29.aspx)|
|[WindowSplit](http://msdn.microsoft.com/library/application-windowsplit-method-project%28Office.15%29.aspx)|
|[WindowUnhide](http://msdn.microsoft.com/library/application-windowunhide-method-project%28Office.15%29.aspx)|
|[WorkOffline](http://msdn.microsoft.com/library/application-workoffline-method-project%28Office.15%29.aspx)|
|[WrapText](http://msdn.microsoft.com/library/application-wraptext-method-project%28Office.15%29.aspx)|
|[Zoom](http://msdn.microsoft.com/library/application-zoom-method-project%28Office.15%29.aspx)|
|[ZoomCalendar](http://msdn.microsoft.com/library/application-zoomcalendar-method-project%28Office.15%29.aspx)|
|[ZoomIn](http://msdn.microsoft.com/library/application-zoomin-method-project%28Office.15%29.aspx)|
|[ZoomOut](http://msdn.microsoft.com/library/application-zoomout-method-project%28Office.15%29.aspx)|
|[ZoomReport](http://msdn.microsoft.com/library/application-zoomreport-method-project%28Office.15%29.aspx)|
|[ZoomTimescale](http://msdn.microsoft.com/library/application-zoomtimescale-method-project%28Office.15%29.aspx)|
|[AddEngagement](http://msdn.microsoft.com/library/application-addengagement-method-project%28Office.15%29.aspx)|
|[EngagementInfo](http://msdn.microsoft.com/library/application-engagementinfo-method-project%28Office.15%29.aspx)|
|[GetDpiScaleFactor](http://msdn.microsoft.com/library/application-getdpiscalefactor-method-project%28Office.15%29.aspx)|
|[InsertTimelineBar](http://msdn.microsoft.com/library/application-addtimelinebar-method-project%28Office.15%29.aspx)|
|[Inspector](http://msdn.microsoft.com/library/application-inspector-method-project%28Office.15%29.aspx)|
|[LocaleName](http://msdn.microsoft.com/library/application-localename-method-project%28Office.15%29.aspx)|
|[ProjectSummaryInfoEx](http://msdn.microsoft.com/library/application-projectsummaryinfoex-method-project%28Office.15%29.aspx)|
|[RefreshEngagementsForProject](http://msdn.microsoft.com/library/application-refreshengagementsforproject-method-project%28Office.15%29.aspx)|
|[RemoveTimelineBar](http://msdn.microsoft.com/library/application-removetimelinebar-method-project%28Office.15%29.aspx)|
|[SubmitAllEngagementsForProject](http://msdn.microsoft.com/library/application-submitallengagementsforproject-method-project%28Office.15%29.aspx)|
|[SubmitSelectedEngagementsForProject](http://msdn.microsoft.com/library/application-submitselectedengagementsforproject-method-project%28Office.15%29.aspx)|
|[TaskOnTimelineEx](http://msdn.microsoft.com/library/application-taskontimelineex-method-project%28Office.15%29.aspx)|
|[TimelineBarDateRange](http://msdn.microsoft.com/library/application-timelinebardaterange-method-project%28Office.15%29.aspx)|
|[UpdateEngagementsForProject](http://msdn.microsoft.com/library/application-updateengagementsforproject-method-project%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[ActiveCell](http://msdn.microsoft.com/library/application-activecell-property-project%28Office.15%29.aspx)|
|[ActiveProject](http://msdn.microsoft.com/library/application-activeproject-property-project%28Office.15%29.aspx)|
|[ActiveSelection](http://msdn.microsoft.com/library/application-activeselection-property-project%28Office.15%29.aspx)|
|[ActiveWindow](http://msdn.microsoft.com/library/application-activewindow-property-project%28Office.15%29.aspx)|
|[AMText](http://msdn.microsoft.com/library/application-amtext-property-project%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/application-application-property-project%28Office.15%29.aspx)|
|[AskToUpdateLinks](http://msdn.microsoft.com/library/application-asktoupdatelinks-property-project%28Office.15%29.aspx)|
|[Assistance](http://msdn.microsoft.com/library/application-assistance-property-project%28Office.15%29.aspx)|
|[AutoClearLeveling](http://msdn.microsoft.com/library/application-autoclearleveling-property-project%28Office.15%29.aspx)|
|[AutoLevel](http://msdn.microsoft.com/library/application-autolevel-property-project%28Office.15%29.aspx)|
|[AutomaticallyFillPhoneticFields](http://msdn.microsoft.com/library/application-automaticallyfillphoneticfields-property-project%28Office.15%29.aspx)|
|[AutomationSecurity](http://msdn.microsoft.com/library/application-automationsecurity-property-project%28Office.15%29.aspx)|
|[Build](http://msdn.microsoft.com/library/application-build-property-project%28Office.15%29.aspx)|
|[Calculation](http://msdn.microsoft.com/library/application-calculation-property-project%28Office.15%29.aspx)|
|[Caption](http://msdn.microsoft.com/library/application-caption-property-project%28Office.15%29.aspx)|
|[CellDragAndDrop](http://msdn.microsoft.com/library/application-celldraganddrop-property-project%28Office.15%29.aspx)|
|[COMAddIns](http://msdn.microsoft.com/library/application-comaddins-property-project%28Office.15%29.aspx)|
|[CommandBars](http://msdn.microsoft.com/library/application-commandbars-property-project%28Office.15%29.aspx)|
|[CompareProjectsCurrentVersionName](http://msdn.microsoft.com/library/application-compareprojectscurrentversionname-property-project%28Office.15%29.aspx)|
|[CompareProjectsPreviousVersionName](http://msdn.microsoft.com/library/application-compareprojectspreviousversionname-property-project%28Office.15%29.aspx)|
|[DateOrder](http://msdn.microsoft.com/library/application-dateorder-property-project%28Office.15%29.aspx)|
|[DateSeparator](http://msdn.microsoft.com/library/application-dateseparator-property-project%28Office.15%29.aspx)|
|[DayLeadingZero](http://msdn.microsoft.com/library/application-dayleadingzero-property-project%28Office.15%29.aspx)|
|[DecimalSeparator](http://msdn.microsoft.com/library/application-decimalseparator-property-project%28Office.15%29.aspx)|
|[DefaultAutoFilter](http://msdn.microsoft.com/library/application-defaultautofilter-property-project%28Office.15%29.aspx)|
|[DefaultDateFormat](http://msdn.microsoft.com/library/application-defaultdateformat-property-project%28Office.15%29.aspx)|
|[DefaultView](http://msdn.microsoft.com/library/application-defaultview-property-project%28Office.15%29.aspx)|
|[DisplayAlerts](http://msdn.microsoft.com/library/application-displayalerts-property-project%28Office.15%29.aspx)|
|[DisplayEntryBar](http://msdn.microsoft.com/library/application-displayentrybar-property-project%28Office.15%29.aspx)|
|[DisplayOLEIndicator](http://msdn.microsoft.com/library/application-displayoleindicator-property-project%28Office.15%29.aspx)|
|[DisplayPlanningWizard](http://msdn.microsoft.com/library/application-displayplanningwizard-property-project%28Office.15%29.aspx)|
|[DisplayProjectGuide](http://msdn.microsoft.com/library/application-displayprojectguide-property-project%28Office.15%29.aspx)|
|[DisplayRecentFiles](http://msdn.microsoft.com/library/application-displayrecentfiles-property-project%28Office.15%29.aspx)|
|[DisplayScheduleMessages](http://msdn.microsoft.com/library/application-displayschedulemessages-property-project%28Office.15%29.aspx)|
|[DisplayScrollBars](http://msdn.microsoft.com/library/application-displayscrollbars-property-project%28Office.15%29.aspx)|
|[DisplayStatusBar](http://msdn.microsoft.com/library/application-displaystatusbar-property-project%28Office.15%29.aspx)|
|[DisplayViewBar](http://msdn.microsoft.com/library/application-displayviewbar-property-project%28Office.15%29.aspx)|
|[DisplayWindowsInTaskbar](http://msdn.microsoft.com/library/application-displaywindowsintaskbar-property-project%28Office.15%29.aspx)|
|[DisplayWizardErrors](http://msdn.microsoft.com/library/application-displaywizarderrors-property-project%28Office.15%29.aspx)|
|[DisplayWizardScheduling](http://msdn.microsoft.com/library/application-displaywizardscheduling-property-project%28Office.15%29.aspx)|
|[DisplayWizardUsage](http://msdn.microsoft.com/library/application-displaywizardusage-property-project%28Office.15%29.aspx)|
|[Edition](http://msdn.microsoft.com/library/application-edition-property-project%28Office.15%29.aspx)|
|[EnableCancelKey](http://msdn.microsoft.com/library/application-enablecancelkey-property-project%28Office.15%29.aspx)|
|[EnableChangeHighlighting](http://msdn.microsoft.com/library/application-enablechangehighlighting-property-project%28Office.15%29.aspx)|
|[EnterpriseAllowLocalBaseCalendars](http://msdn.microsoft.com/library/application-enterpriseallowlocalbasecalendars-property-project%28Office.15%29.aspx)|
|[EnterpriseListSeparator](http://msdn.microsoft.com/library/application-enterpriselistseparator-property-project%28Office.15%29.aspx)|
|[EnterpriseProtectActuals](http://msdn.microsoft.com/library/application-enterpriseprotectactuals-property-project%28Office.15%29.aspx)|
|[FileBuildID](http://msdn.microsoft.com/library/application-filebuildid-property-project%28Office.15%29.aspx)|
|[FileFormatID](http://msdn.microsoft.com/library/application-fileformatid-property-project%28Office.15%29.aspx)|
|[GetCacheStatusForProject](http://msdn.microsoft.com/library/application-getcachestatusforproject-property-project%28Office.15%29.aspx)|
|[GlobalBaseCalendars](http://msdn.microsoft.com/library/application-globalbasecalendars-property-project%28Office.15%29.aspx)|
|[GlobalOutlineCodes](http://msdn.microsoft.com/library/application-globaloutlinecodes-property-project%28Office.15%29.aspx)|
|[GlobalReports](http://msdn.microsoft.com/library/application-globalreports-property-project%28Office.15%29.aspx)|
|[GlobalResourceFilters](http://msdn.microsoft.com/library/application-globalresourcefilters-property-project%28Office.15%29.aspx)|
|[GlobalResourceTables](http://msdn.microsoft.com/library/application-globalresourcetables-property-project%28Office.15%29.aspx)|
|[GlobalTaskFilters](http://msdn.microsoft.com/library/application-globaltaskfilters-property-project%28Office.15%29.aspx)|
|[GlobalTaskTables](http://msdn.microsoft.com/library/application-globaltasktables-property-project%28Office.15%29.aspx)|
|[GlobalViews](http://msdn.microsoft.com/library/application-globalviews-property-project%28Office.15%29.aspx)|
|[GlobalViewsCombination](http://msdn.microsoft.com/library/application-globalviewscombination-property-project%28Office.15%29.aspx)|
|[GlobalViewsSingle](http://msdn.microsoft.com/library/application-globalviewssingle-property-project%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/application-height-property-project%28Office.15%29.aspx)|
|[IsCheckedOut](http://msdn.microsoft.com/library/application-ischeckedout-property-project%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/application-left-property-project%28Office.15%29.aspx)|
|[LevelFreeformTasks](http://msdn.microsoft.com/library/application-levelfreeformtasks-property-project%28Office.15%29.aspx)|
|[LevelIndividualAssignments](http://msdn.microsoft.com/library/application-levelindividualassignments-property-project%28Office.15%29.aspx)|
|[LevelingCanSplit](http://msdn.microsoft.com/library/application-levelingcansplit-property-project%28Office.15%29.aspx)|
|[LevelOrder](http://msdn.microsoft.com/library/application-levelorder-property-project%28Office.15%29.aspx)|
|[LevelPeriodBasis](http://msdn.microsoft.com/library/application-levelperiodbasis-property-project%28Office.15%29.aspx)|
|[LevelProposedBookings](http://msdn.microsoft.com/library/application-levelproposedbookings-property-project%28Office.15%29.aspx)|
|[LevelWithinSlack](http://msdn.microsoft.com/library/application-levelwithinslack-property-project%28Office.15%29.aspx)|
|[ListSeparator](http://msdn.microsoft.com/library/application-listseparator-property-project%28Office.15%29.aspx)|
|[LoadLastFile](http://msdn.microsoft.com/library/application-loadlastfile-property-project%28Office.15%29.aspx)|
|[MonthLeadingZero](http://msdn.microsoft.com/library/application-monthleadingzero-property-project%28Office.15%29.aspx)|
|[MoveAfterReturn](http://msdn.microsoft.com/library/application-moveafterreturn-property-project%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/application-name-property-project%28Office.15%29.aspx)|
|[NewTasksEstimated](http://msdn.microsoft.com/library/application-newtasksestimated-property-project%28Office.15%29.aspx)|
|[OperatingSystem](http://msdn.microsoft.com/library/application-operatingsystem-property-project%28Office.15%29.aspx)|
|[PanZoomFinish](http://msdn.microsoft.com/library/application-panzoomfinish-property-project%28Office.15%29.aspx)|
|[PanZoomStart](http://msdn.microsoft.com/library/application-panzoomstart-property-project%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/application-parent-property-project%28Office.15%29.aspx)|
|[Path](http://msdn.microsoft.com/library/application-path-property-project%28Office.15%29.aspx)|
|[PathSeparator](http://msdn.microsoft.com/library/application-pathseparator-property-project%28Office.15%29.aspx)|
|[PMText](http://msdn.microsoft.com/library/application-pmtext-property-project%28Office.15%29.aspx)|
|[Profiles](http://msdn.microsoft.com/library/application-profiles-property-project%28Office.15%29.aspx)|
|[Projects](http://msdn.microsoft.com/library/application-projects-property-project%28Office.15%29.aspx)|
|[PromptForSummaryInfo](http://msdn.microsoft.com/library/application-promptforsummaryinfo-property-project%28Office.15%29.aspx)|
|[RecentFilesMaximum](http://msdn.microsoft.com/library/application-recentfilesmaximum-property-project%28Office.15%29.aspx)|
|[ScreenUpdating](http://msdn.microsoft.com/library/application-screenupdating-property-project%28Office.15%29.aspx)|
|[ShowAssignmentUnitsAs](http://msdn.microsoft.com/library/application-showassignmentunitsas-property-project%28Office.15%29.aspx)|
|[ShowEstimatedDuration](http://msdn.microsoft.com/library/application-showestimatedduration-property-project%28Office.15%29.aspx)|
|[ShowWelcome](http://msdn.microsoft.com/library/application-showwelcome-property-project%28Office.15%29.aspx)|
|[StartWeekOn](http://msdn.microsoft.com/library/application-startweekon-property-project%28Office.15%29.aspx)|
|[StartYearIn](http://msdn.microsoft.com/library/application-startyearin-property-project%28Office.15%29.aspx)|
|[StatusBar](http://msdn.microsoft.com/library/application-statusbar-property-project%28Office.15%29.aspx)|
|[SupportsMultipleDocuments](http://msdn.microsoft.com/library/application-supportsmultipledocuments-property-project%28Office.15%29.aspx)|
|[SupportsMultipleWindows](http://msdn.microsoft.com/library/application-supportsmultiplewindows-property-project%28Office.15%29.aspx)|
|[ThousandSeparator](http://msdn.microsoft.com/library/application-thousandseparator-property-project%28Office.15%29.aspx)|
|[TimeLeadingZero](http://msdn.microsoft.com/library/application-timeleadingzero-property-project%28Office.15%29.aspx)|
|[TimescaleFinish](http://msdn.microsoft.com/library/application-timescalefinish-property-project%28Office.15%29.aspx)|
|[TimescaleStart](http://msdn.microsoft.com/library/application-timescalestart-property-project%28Office.15%29.aspx)|
|[TimeSeparator](http://msdn.microsoft.com/library/application-timeseparator-property-project%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/application-top-property-project%28Office.15%29.aspx)|
|[TrustProjectServerAndWSSPages](http://msdn.microsoft.com/library/application-trustprojectserverandwsspages-property-project%28Office.15%29.aspx)|
|[TwelveHourTimeFormat](http://msdn.microsoft.com/library/application-twelvehourtimeformat-property-project%28Office.15%29.aspx)|
|[UndoLevels](http://msdn.microsoft.com/library/application-undolevels-property-project%28Office.15%29.aspx)|
|[UsableHeight](http://msdn.microsoft.com/library/application-usableheight-property-project%28Office.15%29.aspx)|
|[UsableWidth](http://msdn.microsoft.com/library/application-usablewidth-property-project%28Office.15%29.aspx)|
|[Use3DLook](http://msdn.microsoft.com/library/application-use3dlook-property-project%28Office.15%29.aspx)|
|[UseOMIDs](http://msdn.microsoft.com/library/application-useomids-property-project%28Office.15%29.aspx)|
|[UserControl](http://msdn.microsoft.com/library/application-usercontrol-property-project%28Office.15%29.aspx)|
|[UserName](http://msdn.microsoft.com/library/application-username-property-project%28Office.15%29.aspx)|
|[VBE](http://msdn.microsoft.com/library/application-vbe-property-project%28Office.15%29.aspx)|
|[Version](http://msdn.microsoft.com/library/application-version-property-project%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/application-visible-property-project%28Office.15%29.aspx)|
|[VisualReportsAdditionalTemplatePath](http://msdn.microsoft.com/library/application-visualreportsadditionaltemplatepath-property-project%28Office.15%29.aspx)|
|[VisualReportTemplateList](http://msdn.microsoft.com/library/application-visualreporttemplatelist-property-project%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/application-width-property-project%28Office.15%29.aspx)|
|[Windows](http://msdn.microsoft.com/library/application-windows-property-project%28Office.15%29.aspx)|
|[Windows2](http://msdn.microsoft.com/library/application-windows2-property-project%28Office.15%29.aspx)|
|[WindowState](http://msdn.microsoft.com/library/application-windowstate-property-project%28Office.15%29.aspx)|

