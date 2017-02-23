---
title: Application Events (Project)
ms.prod: PROJECTSERVER
ms.assetid: 01232fba-5bc2-4c79-9e4d-01a9ec5523da
---


# Application Events (Project)
This object has the following events:

## Events



|**Name**|**Description**|
|:-----|:-----|
|[AfterCubeBuilt](application-aftercubebuilt-event-project.md)|Occurs when the OLAP cube-building process completes.|
|[ApplicationBeforeClose](application-applicationbeforeclose-event-project.md)|Occurs before Project exits.|
|[ConnectionStatusChanged](application-connectionstatuschanged-event-project.md)|Occurs when the status of the connection with Project Server changes. Available only in Project Professional.|
|[IsFunctionalitySupported](application-isfunctionalitysupported-event-project.md)|Occurs after the  **LoadWebBrowserControl** method is called with the third parameter ( _FunctionalityName_) set.|
|[JobCompleted](application-jobcompleted-event-project.md)|Occurs when a queued job originating from Project Professional is completed.|
|[JobStart](application-jobstart-event-project.md)|Occurs before the queue job is put on the server queue. Project Professional only.|
|[LoadWebPage](application-loadwebpage-event-project.md)|Occurs after the  **LoadWebBrowserControl** method is called. The method loads the Web browser control inside Project, and then the event is fired.|
|[LoadWebPane](application-loadwebpane-event-project.md)|Occurs when Project loads a Web pane for  **Task Drivers**,  **Deliverables**, or the  **Project/Resource Import Wizard**.|
|[NewProject](application-newproject-event-project.md)|Occurs when a new project is created, including the default project that is created each time Project starts.|
|[OnUndoOrRedo](application-onundoorredo-event-project.md)|Occurs when a transaction is undone or redone.|
|[PaneActivate](application-paneactivate-event-project.md)|Occurs when the pane is activated.|
|[ProjectAfterSave](application-projectaftersave-event-project.md)|Occurs after a project has been saved.|
|[ProjectAssignmentNew](application-projectassignmentnew-event-project.md)|Occurs when a new assignment is created.|
|[ProjectBeforeAssignmentChange](application-projectbeforeassignmentchange-event-project.md)|Occurs before the user changes the value of an assignment field.|
|[ProjectBeforeAssignmentChange2](application-projectbeforeassignmentchange2-event-project.md)|Occurs before the user changes the value of an assignment field. Uses the  **EventInfo** object parameter.|
|[ProjectBeforeAssignmentDelete](application-projectbeforeassignmentdelete-event-project.md)|Occurs before an assignment is removed or replaced.|
|[ProjectBeforeAssignmentDelete2](application-projectbeforeassignmentdelete2-event-project.md)|Occurs before an assignment is removed or replaced. Uses the  **EventInfo** object parameter.|
|[ProjectBeforeAssignmentNew](application-projectbeforeassignmentnew-event-project.md)|Occurs before one or more assignments are created.|
|[ProjectBeforeAssignmentNew2](application-projectbeforeassignmentnew2-event-project.md)|Occurs before one or more assignments are created. Uses the  **EventInfo** object parameter.|
|[ProjectBeforeClearBaseline](application-projectbeforeclearbaseline-event-project.md)|Occurs before a baseline is cleared. Uses the  **EventInfo** object parameter.|
|[ProjectBeforeClose](application-projectbeforeclose-event-project.md)|Occurs before a project is closed.|
|[ProjectBeforeClose2](application-projectbeforeclose2-event-project.md)|Occurs before a project is closed. Uses the  **EventInfo** object parameter.|
|[ProjectBeforePrint](application-projectbeforeprint-event-project.md)|Occurs before a project is printed.|
|[ProjectBeforePrint2](application-projectbeforeprint2-event-project.md)|Occurs before a project is printed. Uses the  **EventInfo** object parameter.|
|[ProjectBeforePublish](application-projectbeforepublish-event-project.md)|Occurs before a  **Publish** operation is placed on the server queue. The **ProjectBeforePublish** event can be cancelled. Project Professional only.|
|[ProjectBeforeResourceChange](application-projectbeforeresourcechange-event-project.md)|Occurs before the user changes the value of a resource field.|
|[ProjectBeforeResourceChange2](application-projectbeforeresourcechange2-event-project.md)|Occurs before the user changes the value of a resource field. Uses the  **EventInfo** object parameter.|
|[ProjectBeforeResourceDelete](application-projectbeforeresourcedelete-event-project.md)|Occurs before a resource is deleted.|
|[ProjectBeforeResourceDelete2](application-projectbeforeresourcedelete2-event-project.md)|Occurs before a resource is deleted. Uses the  **EventInfo** object parameter.|
|[ProjectBeforeResourceNew](application-projectbeforeresourcenew-event-project.md)|Occurs before one or more resources are created.|
|[ProjectBeforeResourceNew2](application-projectbeforeresourcenew2-event-project.md)|Occurs before one or more resources are created. Uses the  **EventInfo** object parameter.|
|[ProjectBeforeSave](application-projectbeforesave-event-project.md)|Occurs before a project is saved.|
|[ProjectBeforeSave2](application-projectbeforesave2-event-project.md)|Occurs before a project is saved. Uses the  **EventInfo** object parameter.|
|[ProjectBeforeSaveBaseline](application-projectbeforesavebaseline-event-project.md)|Occurs before a baseline is saved. Uses the  **EventInfo** object parameter.|
|[ProjectBeforeTaskChange](application-projectbeforetaskchange-event-project.md)|Occurs before the user changes the value of a task field.|
|[ProjectBeforeTaskChange2](application-projectbeforetaskchange2-event-project.md)|Occurs before the user changes the value of a task field. Uses the  **EventInfo** object parameter.|
|[ProjectBeforeTaskDelete](application-projectbeforetaskdelete-event-project.md)|Occurs before a task is deleted.|
|[ProjectBeforeTaskDelete2](application-projectbeforetaskdelete2-event-project.md)|Occurs before a task is deleted. Uses the  **EventInfo** object parameter.|
|[ProjectBeforeTaskNew](application-projectbeforetasknew-event-project.md)|Occurs before one or more tasks are created.|
|[ProjectBeforeTaskNew2](application-projectbeforetasknew2-event-project.md)|Occurs before one or more tasks are created. Uses the  **EventInfo** object parameter.|
|[ProjectCalculate](application-projectcalculate-event-project.md)|Occurs after a project is calculated.|
|[ProjectResourceNew](application-projectresourcenew-event-project.md)|Occurs before one or more resources is created.|
|[ProjectTaskNew](application-projecttasknew-event-project.md)|Occurs when a new task is created.|
|[SaveCompletedToServer](application-savecompletedtoserver-event-project.md)|Occurs when Project Professional successfully puts the  **Project Save** job in the Project Server Queue.|
|[SaveStartingToServer](application-savestartingtoserver-event-project.md)|Occurs when Project Professional starts to save project changes to the Project Server queue. |
|[SecondaryViewChange](application-secondaryviewchange-event-project.md)|Event occurs when a secondary view pane changes within a project window.|
|[WindowActivate](application-windowactivate-event-project.md)|Occurs when any window within Project is activated. The  **WindowActivate** event does not occur when the application window is activated.|
|[WindowBeforeViewChange](application-windowbeforeviewchange-event-project.md)|Occurs when the top pane view is changed within a window in Project.|
|[WindowDeactivate](application-windowdeactivate-event-project.md)|Occurs when any window within Project is deactivated. The  **WindowDeactivate** event does not occur when the application window is deactivated.|
|[WindowGoalAreaChange](application-windowgoalareachange-event-project.md)|Occurs after a user clicks a different goal area in the Project Guide.|
|[WindowSelectionChange](application-windowselectionchange-event-project.md)|Occurs when the selection handle is changed within a window in Project.|
|[WindowSidepaneDisplayChange](application-windowsidepanedisplaychange-event-project.md)|Occurs when the user shows or hides the Project Guide.|
|[WindowSidepaneTaskChange](application-windowsidepanetaskchange-event-project.md)|Occurs when a user selects different items in the  **Next Steps and Related Activities** menu in the Project Guide.|
|[WindowViewChange](application-windowviewchange-event-project.md)|Occurs after the top pane view is changed within a project window. The  **WindowViewChange** event returns a success argument that tells whether the view change action was successful.|
|[WorkpaneDisplayChange](application-workpanedisplaychange-event-project.md)|Occurs when the Project Guide is hidden or shown.|

