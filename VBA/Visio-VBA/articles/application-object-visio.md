---
title: Application Object (Visio)
keywords: vis_sdr.chm10040
f1_keywords:
- vis_sdr.chm10040
ms.prod: VISIO
api_name:
- Visio.Application
ms.assetid: 5b3c8939-793f-116f-11b8-1d4170d95a63
---


# Application Object (Visio)

Represents an instance of Visio. An external program typically creates or retrieves an  **Application** object before it can retrieve other Visio objects from that instance. Use the Microsoft Visual Basic **CreateObject** function or the **New** keyword to run a new instance, or use the **GetObject** function to retrieve an instance that is already running. You can also use the **CreateObject** function with the **InvisibleApp** object to run a new instance that is invisible. Set the value of the **InvisibleApp** object's **Visible** property to **True** to show it.


## Remarks

Use the  **Documents**, **Windows**, and **Addons** properties of an **Application** object to retrieve the **Document**, **Window**, and **Addon** collections of the instance.

Use the  **ActiveDocument**, **ActivePage**, or **ActiveWindow** property to retrieve the currently active **Document**, **Page**, or **Window** object.


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

Use the  **BuiltInMenus**, **BuiltInToolbars**, **CustomMenus**, **CustomToolbars**, or **CommandBars** property to access the **Application** object's menus and toolbars.

 **ActiveDocument** is the default property of an **Application** object.


 **Note**  Code in the Microsoft Visual Basic for Applications project of a Visio document can use the Visio global object instead of a Visio  **Application** object to retrieve other objects.

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this object maps to the following types:


-  **Microsoft.Office.Interop.Visio.ApplicationClass** (to access the **Application** object.)
    
-  **Microsoft.Office.Interop.Visio.ApplicationClass.Application** (to construct the **Application** object.)
    
-  **Microsoft.Office.Interop.Visio.EApplication_Event** (to access events on the **Application** object.
    

## Events



|**Name**|
|:-----|
|[AfterModal](http://msdn.microsoft.com/library/application-aftermodal-event-visio%28Office.15%29.aspx)|
|[AfterRemoveHiddenInformation](http://msdn.microsoft.com/library/application-afterremovehiddeninformation-event-visio%28Office.15%29.aspx)|
|[AfterReplaceShapes](http://msdn.microsoft.com/library/application-afterreplaceshapes-event-visio%28Office.15%29.aspx)|
|[AfterResume](http://msdn.microsoft.com/library/application-afterresume-event-visio%28Office.15%29.aspx)|
|[AfterResumeEvents](http://msdn.microsoft.com/library/application-afterresumeevents-event-visio%28Office.15%29.aspx)|
|[AppActivated](http://msdn.microsoft.com/library/application-appactivated-event-visio%28Office.15%29.aspx)|
|[AppDeactivated](http://msdn.microsoft.com/library/application-appdeactivated-event-visio%28Office.15%29.aspx)|
|[AppObjActivated](http://msdn.microsoft.com/library/application-appobjactivated-event-visio%28Office.15%29.aspx)|
|[AppObjDeactivated](http://msdn.microsoft.com/library/application-appobjdeactivated-event-visio%28Office.15%29.aspx)|
|[BeforeDataRecordsetDelete](http://msdn.microsoft.com/library/application-beforedatarecordsetdelete-event-visio%28Office.15%29.aspx)|
|[BeforeDocumentClose](http://msdn.microsoft.com/library/application-beforedocumentclose-event-visio%28Office.15%29.aspx)|
|[BeforeDocumentSave](http://msdn.microsoft.com/library/application-beforedocumentsave-event-visio%28Office.15%29.aspx)|
|[BeforeDocumentSaveAs](http://msdn.microsoft.com/library/application-beforedocumentsaveas-event-visio%28Office.15%29.aspx)|
|[BeforeMasterDelete](http://msdn.microsoft.com/library/application-beforemasterdelete-event-visio%28Office.15%29.aspx)|
|[BeforeModal](http://msdn.microsoft.com/library/application-beforemodal-event-visio%28Office.15%29.aspx)|
|[BeforePageDelete](http://msdn.microsoft.com/library/application-beforepagedelete-event-visio%28Office.15%29.aspx)|
|[BeforeQuit](http://msdn.microsoft.com/library/application-beforequit-event-visio%28Office.15%29.aspx)|
|[BeforeReplaceShapes](http://msdn.microsoft.com/library/application-beforereplaceshapes-event-visio%28Office.15%29.aspx)|
|[BeforeSelectionDelete](http://msdn.microsoft.com/library/application-beforeselectiondelete-event-visio%28Office.15%29.aspx)|
|[BeforeShapeDelete](http://msdn.microsoft.com/library/application-beforeshapedelete-event-visio%28Office.15%29.aspx)|
|[BeforeShapeTextEdit](http://msdn.microsoft.com/library/application-beforeshapetextedit-event-visio%28Office.15%29.aspx)|
|[BeforeStyleDelete](http://msdn.microsoft.com/library/application-beforestyledelete-event-visio%28Office.15%29.aspx)|
|[BeforeSuspend](http://msdn.microsoft.com/library/application-beforesuspend-event-visio%28Office.15%29.aspx)|
|[BeforeSuspendEvents](http://msdn.microsoft.com/library/application-beforesuspendevents-event-visio%28Office.15%29.aspx)|
|[BeforeWindowClosed](http://msdn.microsoft.com/library/application-beforewindowclosed-event-visio%28Office.15%29.aspx)|
|[BeforeWindowPageTurn](http://msdn.microsoft.com/library/application-beforewindowpageturn-event-visio%28Office.15%29.aspx)|
|[BeforeWindowSelDelete](http://msdn.microsoft.com/library/application-beforewindowseldelete-event-visio%28Office.15%29.aspx)|
|[CalloutRelationshipAdded](http://msdn.microsoft.com/library/application-calloutrelationshipadded-event-visio%28Office.15%29.aspx)|
|[CalloutRelationshipDeleted](http://msdn.microsoft.com/library/application-calloutrelationshipdeleted-event-visio%28Office.15%29.aspx)|
|[CellChanged](http://msdn.microsoft.com/library/application-cellchanged-event-visio%28Office.15%29.aspx)|
|[ConnectionsAdded](http://msdn.microsoft.com/library/application-connectionsadded-event-visio%28Office.15%29.aspx)|
|[ConnectionsDeleted](http://msdn.microsoft.com/library/application-connectionsdeleted-event-visio%28Office.15%29.aspx)|
|[ContainerRelationshipAdded](http://msdn.microsoft.com/library/application-containerrelationshipadded-event-visio%28Office.15%29.aspx)|
|[ContainerRelationshipDeleted](http://msdn.microsoft.com/library/application-containerrelationshipdeleted-event-visio%28Office.15%29.aspx)|
|[ConvertToGroupCanceled](http://msdn.microsoft.com/library/application-converttogroupcanceled-event-visio%28Office.15%29.aspx)|
|[DataRecordsetAdded](http://msdn.microsoft.com/library/application-datarecordsetadded-event-visio%28Office.15%29.aspx)|
|[DataRecordsetChanged](http://msdn.microsoft.com/library/application-datarecordsetchanged-event-visio%28Office.15%29.aspx)|
|[DesignModeEntered](http://msdn.microsoft.com/library/application-designmodeentered-event-visio%28Office.15%29.aspx)|
|[DocumentChanged](http://msdn.microsoft.com/library/application-documentchanged-event-visio%28Office.15%29.aspx)|
|[DocumentCloseCanceled](http://msdn.microsoft.com/library/application-documentclosecanceled-event-visio%28Office.15%29.aspx)|
|[DocumentCreated](http://msdn.microsoft.com/library/application-documentcreated-event-visio%28Office.15%29.aspx)|
|[DocumentOpened](http://msdn.microsoft.com/library/application-documentopened-event-visio%28Office.15%29.aspx)|
|[DocumentSaved](http://msdn.microsoft.com/library/application-documentsaved-event-visio%28Office.15%29.aspx)|
|[DocumentSavedAs](http://msdn.microsoft.com/library/application-documentsavedas-event-visio%28Office.15%29.aspx)|
|[EnterScope](http://msdn.microsoft.com/library/application-enterscope-event-visio%28Office.15%29.aspx)|
|[ExitScope](http://msdn.microsoft.com/library/application-exitscope-event-visio%28Office.15%29.aspx)|
|[FormulaChanged](http://msdn.microsoft.com/library/application-formulachanged-event-visio%28Office.15%29.aspx)|
|[GroupCanceled](http://msdn.microsoft.com/library/application-groupcanceled-event-visio%28Office.15%29.aspx)|
|[KeyDown](http://msdn.microsoft.com/library/application-keydown-event-visio%28Office.15%29.aspx)|
|[KeyPress](http://msdn.microsoft.com/library/application-keypress-event-visio%28Office.15%29.aspx)|
|[KeyUp](http://msdn.microsoft.com/library/application-keyup-event-visio%28Office.15%29.aspx)|
|[MarkerEvent](http://msdn.microsoft.com/library/application-markerevent-event-visio%28Office.15%29.aspx)|
|[MasterAdded](http://msdn.microsoft.com/library/application-masteradded-event-visio%28Office.15%29.aspx)|
|[MasterChanged](http://msdn.microsoft.com/library/application-masterchanged-event-visio%28Office.15%29.aspx)|
|[MasterDeleteCanceled](http://msdn.microsoft.com/library/application-masterdeletecanceled-event-visio%28Office.15%29.aspx)|
|[MouseDown](http://msdn.microsoft.com/library/application-mousedown-event-visio%28Office.15%29.aspx)|
|[MouseMove](http://msdn.microsoft.com/library/application-mousemove-event-visio%28Office.15%29.aspx)|
|[MouseUp](http://msdn.microsoft.com/library/application-mouseup-event-visio%28Office.15%29.aspx)|
|[MustFlushScopeBeginning](http://msdn.microsoft.com/library/application-mustflushscopebeginning-event-visio%28Office.15%29.aspx)|
|[MustFlushScopeEnded](http://msdn.microsoft.com/library/application-mustflushscopeended-event-visio%28Office.15%29.aspx)|
|[NoEventsPending](http://msdn.microsoft.com/library/application-noeventspending-event-visio%28Office.15%29.aspx)|
|[OnKeystrokeMessageForAddon](http://msdn.microsoft.com/library/application-onkeystrokemessageforaddon-event-visio%28Office.15%29.aspx)|
|[PageAdded](http://msdn.microsoft.com/library/application-pageadded-event-visio%28Office.15%29.aspx)|
|[PageChanged](http://msdn.microsoft.com/library/application-pagechanged-event-visio%28Office.15%29.aspx)|
|[PageDeleteCanceled](http://msdn.microsoft.com/library/application-pagedeletecanceled-event-visio%28Office.15%29.aspx)|
|[QueryCancelConvertToGroup](http://msdn.microsoft.com/library/application-querycancelconverttogroup-event-visio%28Office.15%29.aspx)|
|[QueryCancelDocumentClose](http://msdn.microsoft.com/library/application-querycanceldocumentclose-event-visio%28Office.15%29.aspx)|
|[QueryCancelGroup](http://msdn.microsoft.com/library/application-querycancelgroup-event-visio%28Office.15%29.aspx)|
|[QueryCancelMasterDelete](http://msdn.microsoft.com/library/application-querycancelmasterdelete-event-visio%28Office.15%29.aspx)|
|[QueryCancelPageDelete](http://msdn.microsoft.com/library/application-querycancelpagedelete-event-visio%28Office.15%29.aspx)|
|[QueryCancelQuit](http://msdn.microsoft.com/library/application-querycancelquit-event-visio%28Office.15%29.aspx)|
|[QueryCancelReplaceShapes](http://msdn.microsoft.com/library/application-querycancelreplaceshapes-event-visio%28Office.15%29.aspx)|
|[QueryCancelSelectionDelete](http://msdn.microsoft.com/library/application-querycancelselectiondelete-event-visio%28Office.15%29.aspx)|
|[QueryCancelStyleDelete](http://msdn.microsoft.com/library/application-querycancelstyledelete-event-visio%28Office.15%29.aspx)|
|[QueryCancelSuspend](http://msdn.microsoft.com/library/application-querycancelsuspend-event-visio%28Office.15%29.aspx)|
|[QueryCancelSuspendEvents](http://msdn.microsoft.com/library/application-querycancelsuspendevents-event-visio%28Office.15%29.aspx)|
|[QueryCancelUngroup](http://msdn.microsoft.com/library/application-querycancelungroup-event-visio%28Office.15%29.aspx)|
|[QueryCancelWindowClose](http://msdn.microsoft.com/library/application-querycancelwindowclose-event-visio%28Office.15%29.aspx)|
|[QuitCanceled](http://msdn.microsoft.com/library/application-quitcanceled-event-visio%28Office.15%29.aspx)|
|[ReplaceShapesCanceled](http://msdn.microsoft.com/library/application-replaceshapescanceled-event-visio%28Office.15%29.aspx)|
|[RuleSetValidated](http://msdn.microsoft.com/library/application-rulesetvalidated-event-visio%28Office.15%29.aspx)|
|[RunModeEntered](http://msdn.microsoft.com/library/application-runmodeentered-event-visio%28Office.15%29.aspx)|
|[SelectionAdded](http://msdn.microsoft.com/library/application-selectionadded-event-visio%28Office.15%29.aspx)|
|[SelectionChanged](http://msdn.microsoft.com/library/application-selectionchanged-event-visio%28Office.15%29.aspx)|
|[SelectionDeleteCanceled](http://msdn.microsoft.com/library/application-selectiondeletecanceled-event-visio%28Office.15%29.aspx)|
|[ShapeAdded](http://msdn.microsoft.com/library/application-shapeadded-event-visio%28Office.15%29.aspx)|
|[ShapeChanged](http://msdn.microsoft.com/library/application-shapechanged-event-visio%28Office.15%29.aspx)|
|[ShapeDataGraphicChanged](http://msdn.microsoft.com/library/application-shapedatagraphicchanged-event-visio%28Office.15%29.aspx)|
|[ShapeExitedTextEdit](http://msdn.microsoft.com/library/application-shapeexitedtextedit-event-visio%28Office.15%29.aspx)|
|[ShapeLinkAdded](http://msdn.microsoft.com/library/application-shapelinkadded-event-visio%28Office.15%29.aspx)|
|[ShapeLinkDeleted](http://msdn.microsoft.com/library/application-shapelinkdeleted-event-visio%28Office.15%29.aspx)|
|[ShapeParentChanged](http://msdn.microsoft.com/library/application-shapeparentchanged-event-visio%28Office.15%29.aspx)|
|[StyleAdded](http://msdn.microsoft.com/library/application-styleadded-event-visio%28Office.15%29.aspx)|
|[StyleChanged](http://msdn.microsoft.com/library/application-stylechanged-event-visio%28Office.15%29.aspx)|
|[StyleDeleteCanceled](http://msdn.microsoft.com/library/application-styledeletecanceled-event-visio%28Office.15%29.aspx)|
|[SuspendCanceled](http://msdn.microsoft.com/library/application-suspendcanceled-event-visio%28Office.15%29.aspx)|
|[SuspendEventsCanceled](http://msdn.microsoft.com/library/application-suspendeventscanceled-event-visio%28Office.15%29.aspx)|
|[TextChanged](http://msdn.microsoft.com/library/application-textchanged-event-visio%28Office.15%29.aspx)|
|[UngroupCanceled](http://msdn.microsoft.com/library/application-ungroupcanceled-event-visio%28Office.15%29.aspx)|
|[ViewChanged](http://msdn.microsoft.com/library/application-viewchanged-event-visio%28Office.15%29.aspx)|
|[VisioIsIdle](http://msdn.microsoft.com/library/application-visioisidle-event-visio%28Office.15%29.aspx)|
|[WindowActivated](http://msdn.microsoft.com/library/application-windowactivated-event-visio%28Office.15%29.aspx)|
|[WindowChanged](http://msdn.microsoft.com/library/application-windowchanged-event-visio%28Office.15%29.aspx)|
|[WindowCloseCanceled](http://msdn.microsoft.com/library/application-windowclosecanceled-event-visio%28Office.15%29.aspx)|
|[WindowOpened](http://msdn.microsoft.com/library/application-windowopened-event-visio%28Office.15%29.aspx)|
|[WindowTurnedToPage](http://msdn.microsoft.com/library/application-windowturnedtopage-event-visio%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[AddUndoUnit](http://msdn.microsoft.com/library/application-addundounit-method-visio%28Office.15%29.aspx)|
|[BeginUndoScope](http://msdn.microsoft.com/library/application-beginundoscope-method-visio%28Office.15%29.aspx)|
|[ClearCustomMenus](http://msdn.microsoft.com/library/application-clearcustommenus-method-visio%28Office.15%29.aspx)|
|[ClearCustomToolbars](http://msdn.microsoft.com/library/application-clearcustomtoolbars-method-visio%28Office.15%29.aspx)|
|[ConvertResult](http://msdn.microsoft.com/library/application-convertresult-method-visio%28Office.15%29.aspx)|
|[DoCmd](http://msdn.microsoft.com/library/application-docmd-method-visio%28Office.15%29.aspx)|
|[EndUndoScope](http://msdn.microsoft.com/library/application-endundoscope-method-visio%28Office.15%29.aspx)|
|[EnumDirectories](http://msdn.microsoft.com/library/application-enumdirectories-method-visio%28Office.15%29.aspx)|
|[FormatResult](http://msdn.microsoft.com/library/application-formatresult-method-visio%28Office.15%29.aspx)|
|[FormatResultEx](http://msdn.microsoft.com/library/application-formatresultex-method-visio%28Office.15%29.aspx)|
|[GetBuiltInStencilFile](http://msdn.microsoft.com/library/application-getbuiltinstencilfile-method-visio%28Office.15%29.aspx)|
|[GetCustomStencilFile](http://msdn.microsoft.com/library/application-getcustomstencilfile-method-visio%28Office.15%29.aspx)|
|[GetPreviewEnabled](http://msdn.microsoft.com/library/application-getpreviewenabled-method-visio%28Office.15%29.aspx)|
|[InvokeHelp](http://msdn.microsoft.com/library/application-invokehelp-method-visio%28Office.15%29.aspx)|
|[OnComponentEnterState](http://msdn.microsoft.com/library/application-oncomponententerstate-method-visio%28Office.15%29.aspx)|
|[PurgeUndo](http://msdn.microsoft.com/library/application-purgeundo-method-visio%28Office.15%29.aspx)|
|[QueueMarkerEvent](http://msdn.microsoft.com/library/application-queuemarkerevent-method-visio%28Office.15%29.aspx)|
|[Quit](http://msdn.microsoft.com/library/application-quit-method-visio%28Office.15%29.aspx)|
|[Redo](http://msdn.microsoft.com/library/application-redo-method-visio%28Office.15%29.aspx)|
|[RegisterRibbonX](http://msdn.microsoft.com/library/application-registerribbonx-method-visio%28Office.15%29.aspx)|
|[RenameCurrentScope](http://msdn.microsoft.com/library/application-renamecurrentscope-method-visio%28Office.15%29.aspx)|
|[SetCustomMenus](http://msdn.microsoft.com/library/application-setcustommenus-method-visio%28Office.15%29.aspx)|
|[SetCustomToolbars](http://msdn.microsoft.com/library/application-setcustomtoolbars-method-visio%28Office.15%29.aspx)|
|[SetPreviewEnabled](http://msdn.microsoft.com/library/application-setpreviewenabled-method-visio%28Office.15%29.aspx)|
|[Undo](http://msdn.microsoft.com/library/application-undo-method-visio%28Office.15%29.aspx)|
|[UnregisterRibbonX](http://msdn.microsoft.com/library/application-unregisterribbonx-method-visio%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Active](http://msdn.microsoft.com/library/application-active-property-visio%28Office.15%29.aspx)|
|[ActiveDocument](http://msdn.microsoft.com/library/application-activedocument-property-visio%28Office.15%29.aspx)|
|[ActivePage](http://msdn.microsoft.com/library/application-activepage-property-visio%28Office.15%29.aspx)|
|[ActivePrinter](http://msdn.microsoft.com/library/application-activeprinter-property-visio%28Office.15%29.aspx)|
|[ActiveWindow](http://msdn.microsoft.com/library/application-activewindow-property-visio%28Office.15%29.aspx)|
|[AddonPaths](http://msdn.microsoft.com/library/application-addonpaths-property-visio%28Office.15%29.aspx)|
|[Addons](http://msdn.microsoft.com/library/application-addons-property-visio%28Office.15%29.aspx)|
|[AlertResponse](http://msdn.microsoft.com/library/application-alertresponse-property-visio%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/application-application-property-visio%28Office.15%29.aspx)|
|[Assistance](http://msdn.microsoft.com/library/application-assistance-property-visio%28Office.15%29.aspx)|
|[AutoLayout](http://msdn.microsoft.com/library/application-autolayout-property-visio%28Office.15%29.aspx)|
|[AutoRecoverInterval](http://msdn.microsoft.com/library/application-autorecoverinterval-property-visio%28Office.15%29.aspx)|
|[AvailablePrinters](http://msdn.microsoft.com/library/application-availableprinters-property-visio%28Office.15%29.aspx)|
|[Build](http://msdn.microsoft.com/library/application-build-property-visio%28Office.15%29.aspx)|
|[BuiltInMenus](http://msdn.microsoft.com/library/application-builtinmenus-property-visio%28Office.15%29.aspx)|
|[BuiltInToolbars](http://msdn.microsoft.com/library/application-builtintoolbars-property-visio%28Office.15%29.aspx)|
|[COMAddIns](http://msdn.microsoft.com/library/application-comaddins-property-visio%28Office.15%29.aspx)|
|[CommandBars](http://msdn.microsoft.com/library/application-commandbars-property-visio%28Office.15%29.aspx)|
|[CommandLine](http://msdn.microsoft.com/library/application-commandline-property-visio%28Office.15%29.aspx)|
|[ConnectorToolDataObject](http://msdn.microsoft.com/library/application-connectortooldataobject-property-visio%28Office.15%29.aspx)|
|[CurrentEdition](http://msdn.microsoft.com/library/application-currentedition-property-visio%28Office.15%29.aspx)|
|[CurrentScope](http://msdn.microsoft.com/library/application-currentscope-property-visio%28Office.15%29.aspx)|
|[CustomMenus](http://msdn.microsoft.com/library/application-custommenus-property-visio%28Office.15%29.aspx)|
|[CustomMenusFile](http://msdn.microsoft.com/library/application-custommenusfile-property-visio%28Office.15%29.aspx)|
|[CustomToolbars](http://msdn.microsoft.com/library/application-customtoolbars-property-visio%28Office.15%29.aspx)|
|[CustomToolbarsFile](http://msdn.microsoft.com/library/application-customtoolbarsfile-property-visio%28Office.15%29.aspx)|
|[DataFeaturesEnabled](http://msdn.microsoft.com/library/application-datafeaturesenabled-property-visio%28Office.15%29.aspx)|
|[DefaultAngleUnits](http://msdn.microsoft.com/library/application-defaultangleunits-property-visio%28Office.15%29.aspx)|
|[DefaultDurationUnits](http://msdn.microsoft.com/library/application-defaultdurationunits-property-visio%28Office.15%29.aspx)|
|[DefaultRectangleDataObject](http://msdn.microsoft.com/library/application-defaultrectangledataobject-property-visio%28Office.15%29.aspx)|
|[DefaultTextUnits](http://msdn.microsoft.com/library/application-defaulttextunits-property-visio%28Office.15%29.aspx)|
|[DefaultZoomBehavior](http://msdn.microsoft.com/library/application-defaultzoombehavior-property-visio%28Office.15%29.aspx)|
|[DeferRecalc](http://msdn.microsoft.com/library/application-deferrecalc-property-visio%28Office.15%29.aspx)|
|[DeferRelationshipRecalc](http://msdn.microsoft.com/library/application-deferrelationshiprecalc-property-visio%28Office.15%29.aspx)|
|[DialogFont](http://msdn.microsoft.com/library/application-dialogfont-property-visio%28Office.15%29.aspx)|
|[Documents](http://msdn.microsoft.com/library/application-documents-property-visio%28Office.15%29.aspx)|
|[DrawingPaths](http://msdn.microsoft.com/library/application-drawingpaths-property-visio%28Office.15%29.aspx)|
|[EventInfo](http://msdn.microsoft.com/library/application-eventinfo-property-visio%28Office.15%29.aspx)|
|[EventList](http://msdn.microsoft.com/library/application-eventlist-property-visio%28Office.15%29.aspx)|
|[EventsEnabled](http://msdn.microsoft.com/library/application-eventsenabled-property-visio%28Office.15%29.aspx)|
|[FullBuild](http://msdn.microsoft.com/library/application-fullbuild-property-visio%28Office.15%29.aspx)|
|[HelpPaths](http://msdn.microsoft.com/library/application-helppaths-property-visio%28Office.15%29.aspx)|
|[InhibitSelectChange](http://msdn.microsoft.com/library/application-inhibitselectchange-property-visio%28Office.15%29.aspx)|
|[InstanceHandle32](http://msdn.microsoft.com/library/application-instancehandle32-property-visio%28Office.15%29.aspx)|
|[InstanceHandle64](http://msdn.microsoft.com/library/application-instancehandle64-property-visio%28Office.15%29.aspx)|
|[IsInScope](http://msdn.microsoft.com/library/application-isinscope-property-visio%28Office.15%29.aspx)|
|[IsUndoingOrRedoing](http://msdn.microsoft.com/library/application-isundoingorredoing-property-visio%28Office.15%29.aspx)|
|[IsVisio32](http://msdn.microsoft.com/library/application-isvisio32-property-visio%28Office.15%29.aspx)|
|[Language](http://msdn.microsoft.com/library/application-language-property-visio%28Office.15%29.aspx)|
|[LanguageHelp](http://msdn.microsoft.com/library/application-languagehelp-property-visio%28Office.15%29.aspx)|
|[LanguageSettings](http://msdn.microsoft.com/library/application-languagesettings-property-visio%28Office.15%29.aspx)|
|[LiveDynamics](http://msdn.microsoft.com/library/application-livedynamics-property-visio%28Office.15%29.aspx)|
|[MyShapesPath](http://msdn.microsoft.com/library/application-myshapespath-property-visio%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/application-name-property-visio%28Office.15%29.aspx)|
|[ObjectType](http://msdn.microsoft.com/library/application-objecttype-property-visio%28Office.15%29.aspx)|
|[OnDataChangeDelay](http://msdn.microsoft.com/library/application-ondatachangedelay-property-visio%28Office.15%29.aspx)|
|[Path](http://msdn.microsoft.com/library/application-path-property-visio%28Office.15%29.aspx)|
|[PersistsEvents](http://msdn.microsoft.com/library/application-persistsevents-property-visio%28Office.15%29.aspx)|
|[ProcessID](http://msdn.microsoft.com/library/application-processid-property-visio%28Office.15%29.aspx)|
|[PromptForSummary](http://msdn.microsoft.com/library/application-promptforsummary-property-visio%28Office.15%29.aspx)|
|[SaveAsWebObject](http://msdn.microsoft.com/library/application-saveaswebobject-property-visio%28Office.15%29.aspx)|
|[ScreenUpdating](http://msdn.microsoft.com/library/application-screenupdating-property-visio%28Office.15%29.aspx)|
|[Settings](http://msdn.microsoft.com/library/application-settings-property-visio%28Office.15%29.aspx)|
|[ShowChanges](http://msdn.microsoft.com/library/application-showchanges-property-visio%28Office.15%29.aspx)|
|[ShowProgress](http://msdn.microsoft.com/library/application-showprogress-property-visio%28Office.15%29.aspx)|
|[ShowStatusBar](http://msdn.microsoft.com/library/application-showstatusbar-property-visio%28Office.15%29.aspx)|
|[ShowToolbar](http://msdn.microsoft.com/library/application-showtoolbar-property-visio%28Office.15%29.aspx)|
|[StartupPaths](http://msdn.microsoft.com/library/application-startuppaths-property-visio%28Office.15%29.aspx)|
|[Stat](http://msdn.microsoft.com/library/application-stat-property-visio%28Office.15%29.aspx)|
|[StencilPaths](http://msdn.microsoft.com/library/application-stencilpaths-property-visio%28Office.15%29.aspx)|
|[TemplatePaths](http://msdn.microsoft.com/library/application-templatepaths-property-visio%28Office.15%29.aspx)|
|[TraceFlags](http://msdn.microsoft.com/library/application-traceflags-property-visio%28Office.15%29.aspx)|
|[TypelibMajorVersion](http://msdn.microsoft.com/library/application-typelibmajorversion-property-visio%28Office.15%29.aspx)|
|[TypelibMinorVersion](http://msdn.microsoft.com/library/application-typelibminorversion-property-visio%28Office.15%29.aspx)|
|[UndoEnabled](http://msdn.microsoft.com/library/application-undoenabled-property-visio%28Office.15%29.aspx)|
|[UserName](http://msdn.microsoft.com/library/application-username-property-visio%28Office.15%29.aspx)|
|[VBAEnabled](http://msdn.microsoft.com/library/application-vbaenabled-property-visio%28Office.15%29.aspx)|
|[Vbe](http://msdn.microsoft.com/library/application-vbe-property-visio%28Office.15%29.aspx)|
|[Version](http://msdn.microsoft.com/library/application-version-property-visio%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/application-visible-property-visio%28Office.15%29.aspx)|
|[Window](http://msdn.microsoft.com/library/application-window-property-visio%28Office.15%29.aspx)|
|[WindowHandle32](http://msdn.microsoft.com/library/application-windowhandle32-property-visio%28Office.15%29.aspx)|
|[Windows](http://msdn.microsoft.com/library/application-windows-property-visio%28Office.15%29.aspx)|

