---
title: Application Members (Visio)
ms.prod: VISIO
ms.assetid: b0981ef5-7cf5-aee7-3750-36af0cff0c01
---


# Application Members (Visio)
Represents an instance of Visio. An external program typically creates or retrieves an  **Application** object before it can retrieve other Visio objects from that instance. Use the Microsoft Visual Basic **CreateObject** function or the **New** keyword to run a new instance, or use the **GetObject** function to retrieve an instance that is already running. You can also use the **CreateObject** function with the **InvisibleApp** object to run a new instance that is invisible. Set the value of the **InvisibleApp** object's **Visible** property to **True** to show it.

Represents an instance of Visio. An external program typically creates or retrieves an  **Application** object before it can retrieve other Visio objects from that instance. Use the Microsoft Visual Basic **CreateObject** function or the **New** keyword to run a new instance, or use the **GetObject** function to retrieve an instance that is already running. You can also use the **CreateObject** function with the **InvisibleApp** object to run a new instance that is invisible. Set the value of the **InvisibleApp** object's **Visible** property to **True** to show it.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[AfterModal](application-aftermodal-event-visio.md)|Occurs after the Microsoft Visio instance leaves a modal state.|
|[AfterRemoveHiddenInformation](application-afterremovehiddeninformation-event-visio.md)|Occurs when hidden information is removed from the document.|
|[AfterReplaceShapes](application-afterreplaceshapes-event-visio.md)|Occurs after a shape-replacement operation.|
|[AfterResume](application-afterresume-event-visio.md)|Occurs when the operating system resumes normal operation after having been suspended.|
|[AfterResumeEvents](application-afterresumeevents-event-visio.md)|Occurs after firing of events is resumed.|
|[AppActivated](application-appactivated-event-visio.md)|Occurs after a Microsoft Visio instance becomes active.|
|[AppDeactivated](application-appdeactivated-event-visio.md)|Occurs after a Microsoft Visio instance becomes inactive.|
|[AppObjActivated](application-appobjactivated-event-visio.md)|Occurs after a Microsoft Visio instance becomes active.|
|[AppObjDeactivated](application-appobjdeactivated-event-visio.md)|Occurs after a Microsoft Visio instance becomes inactive.|
|[BeforeDataRecordsetDelete](application-beforedatarecordsetdelete-event-visio.md)|Occurs before a  **DataRecordset** object is deleted from the **DataRecordsets** collection.|
|[BeforeDocumentClose](application-beforedocumentclose-event-visio.md)|Occurs before a document is closed.|
|[BeforeDocumentSave](application-beforedocumentsave-event-visio.md)|Occurs before a document is saved.|
|[BeforeDocumentSaveAs](application-beforedocumentsaveas-event-visio.md)|Occurs just before a document is saved by using the  **Save As** command.|
|[BeforeMasterDelete](application-beforemasterdelete-event-visio.md)|Occurs before a master is deleted from a document.|
|[BeforeModal](application-beforemodal-event-visio.md)|Occurs before a Microsoft Visio instance enters a modal state.|
|[BeforePageDelete](application-beforepagedelete-event-visio.md)|Occurs before a page is deleted.|
|[BeforeQuit](application-beforequit-event-visio.md)|Occurs before a Microsoft Visio instance terminates.|
|[BeforeReplaceShapes](application-beforereplaceshapes-event-visio.md)|Occurs just before a shape-replacement operation.|
|[BeforeSelectionDelete](application-beforeselectiondelete-event-visio.md)|Occurs before selected objects are deleted.|
|[BeforeShapeDelete](application-beforeshapedelete-event-visio.md)|Occurs before a shape is deleted.|
|[BeforeShapeTextEdit](application-beforeshapetextedit-event-visio.md)|Occurs before a shape is opened for text editing in the user interface.|
|[BeforeStyleDelete](application-beforestyledelete-event-visio.md)|Occurs before a style is deleted.|
|[BeforeSuspend](application-beforesuspend-event-visio.md)|Occurs before the operating system enters a suspended state.|
|[BeforeSuspendEvents](application-beforesuspendevents-event-visio.md)|Occurs before firing of events is suspended.|
|[BeforeWindowClosed](application-beforewindowclosed-event-visio.md)|Occurs before a window is closed.|
|[BeforeWindowPageTurn](application-beforewindowpageturn-event-visio.md)|Occurs before a window is about to show a different page.|
|[BeforeWindowSelDelete](application-beforewindowseldelete-event-visio.md)|Occurs before the shapes in the selection of a window are deleted.|
|[CalloutRelationshipAdded](application-calloutrelationshipadded-event-visio.md)|Occurs when a new callout relationship is added to the application.|
|[CalloutRelationshipDeleted](application-calloutrelationshipdeleted-event-visio.md)|Occurs when a callout relationship is deleted from the application.|
|[CellChanged](application-cellchanged-event-visio.md)|Occurs after the value changes in a cell in a document.|
|[ConnectionsAdded](application-connectionsadded-event-visio.md)|Occurs after connections have been established between shapes.|
|[ConnectionsDeleted](application-connectionsdeleted-event-visio.md)|Occurs after connections between shapes have been removed.|
|[ContainerRelationshipAdded](application-containerrelationshipadded-event-visio.md)|Occurs when a new container relationship is added to the document.|
|[ContainerRelationshipDeleted](application-containerrelationshipdeleted-event-visio.md)|Occurs when a container relationship is deleted from the document.|
|[ConvertToGroupCanceled](application-converttogroupcanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelConvertToGroup** event.|
|[DataRecordsetAdded](application-datarecordsetadded-event-visio.md)|Occurs when a  **DataRecordset** object is added to a **DataRecordsets** collection.|
|[DataRecordsetChanged](application-datarecordsetchanged-event-visio.md)|Occurs when a data recordset changes as a result of being refreshed.|
|[DesignModeEntered](application-designmodeentered-event-visio.md)|Occurs before a document enters design mode.|
|[DocumentChanged](application-documentchanged-event-visio.md)|Occurs after certain properties of a document are changed.|
|[DocumentCloseCanceled](application-documentclosecanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelDocumentClose** event.|
|[DocumentCreated](application-documentcreated-event-visio.md)|Occurs after a document is created.|
|[DocumentOpened](application-documentopened-event-visio.md)|Occurs after a document is opened.|
|[DocumentSaved](application-documentsaved-event-visio.md)|Occurs after a document is saved.|
|[DocumentSavedAs](application-documentsavedas-event-visio.md)|Occurs after a document is saved by using the  **Save As** command.|
|[EnterScope](application-enterscope-event-visio.md)|Queued when an internal command begins, or when an Automation client opens a scope by using the  **BeginUndoScope** method.|
|[ExitScope](application-exitscope-event-visio.md)|Queued when an internal command ends, or when an Automation client exits a scope by using the  **EndUndoScope** method.|
|[FormulaChanged](application-formulachanged-event-visio.md)|Occurs after a formula changes in a cell in the object that receives the event.|
|[GroupCanceled](application-groupcanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelGroup** event.|
|[KeyDown](application-keydown-event-visio.md)|Occurs when a keyboard key is pressed.|
|[KeyPress](application-keypress-event-visio.md)|Occurs when a keyboard key is pressed.|
|[KeyUp](application-keyup-event-visio.md)|Occurs when a keyboard key is released.|
|[MarkerEvent](application-markerevent-event-visio.md)|Caused by calling the  **QueueMarkerEvent** method.|
|[MasterAdded](application-masteradded-event-visio.md)|Occurs after a new master is added to a document.|
|[MasterChanged](application-masterchanged-event-visio.md)|Occurs after properties of a master are changed and propagated to its instances.|
|[MasterDeleteCanceled](application-masterdeletecanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelMasterDelete** event.|
|[MouseDown](application-mousedown-event-visio.md)|Occurs when a mouse button is clicked.|
|[MouseMove](application-mousemove-event-visio.md)|Occurs when the mouse is moved.|
|[MouseUp](application-mouseup-event-visio.md)|Occurs when a mouse button is released.|
|[MustFlushScopeBeginning](application-mustflushscopebeginning-event-visio.md)|Occurs before the Microsoft Visio instance is forced to flush its event queue.|
|[MustFlushScopeEnded](application-mustflushscopeended-event-visio.md)|Occurs after the Microsoft Visio instance is forced to flush its event queue.|
|[NoEventsPending](application-noeventspending-event-visio.md)|Occurs after the Microsoft Visio instance flushes its event queue.|
|[OnKeystrokeMessageForAddon](application-onkeystrokemessageforaddon-event-visio.md)|Occurs when Microsoft Visio receives a keystroke message from Microsoft Windows that is targeted at an add-on window or child of an add-on window.|
|[PageAdded](application-pageadded-event-visio.md)|Occurs after a new page is added to a document.|
|[PageChanged](application-pagechanged-event-visio.md)|Occurs after the name of a page, the background page associated with a page, or the page type (foreground or background) changes.|
|[PageDeleteCanceled](application-pagedeletecanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelPageDelete** event.|
|[QueryCancelConvertToGroup](application-querycancelconverttogroup-event-visio.md)|Occurs before the application converts a selection of shapes to a group in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelDocumentClose](application-querycanceldocumentclose-event-visio.md)|Occurs before the application closes a document in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelGroup](application-querycancelgroup-event-visio.md)|Occurs before the application groups a selection of shapes in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelMasterDelete](application-querycancelmasterdelete-event-visio.md)|Occurs before the application deletes a master in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelPageDelete](application-querycancelpagedelete-event-visio.md)|Occurs before the application deletes a page in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelQuit](application-querycancelquit-event-visio.md)|Occurs before the application terminates in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelReplaceShapes](application-querycancelreplaceshapes-event-visio.md)|Occurs immediately after a shape-replacement operation is requested. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelSelectionDelete](application-querycancelselectiondelete-event-visio.md)|Occurs before the application deletes a selection of shapes in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelStyleDelete](application-querycancelstyledelete-event-visio.md)|Occurs before the application deletes a style in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelSuspend](application-querycancelsuspend-event-visio.md)|Occurs before the operating system enters a suspended state. If any event handler returns  **True** , the Microsoft Visio instance will deny the operating system's request.|
|[QueryCancelSuspendEvents](application-querycancelsuspendevents-event-visio.md)|Occurs before the application suspends events in response to client code. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelUngroup](application-querycancelungroup-event-visio.md)|Occurs before the application ungroups a selection of shapes in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelWindowClose](application-querycancelwindowclose-event-visio.md)|Occurs before the application closes a window in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[QuitCanceled](application-quitcanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelQuit** event.|
|[ReplaceShapesCanceled](application-replaceshapescanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelReplaceShapes** event.|
|[RuleSetValidated](application-rulesetvalidated-event-visio.md)|Occurs when a rule set is validated.|
|[RunModeEntered](application-runmodeentered-event-visio.md)|Occurs after a document enters run mode.|
|[SelectionAdded](application-selectionadded-event-visio.md)|Occurs after one or more shapes are added to a document.|
|[SelectionChanged](application-selectionchanged-event-visio.md)|Occurs after a set of shapes selected in a window changes.|
|[SelectionDeleteCanceled](application-selectiondeletecanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelSelectionDelete** event.|
|[ShapeAdded](application-shapeadded-event-visio.md)|Occurs after one or more shapes are added to a document.|
|[ShapeChanged](application-shapechanged-event-visio.md)|Occurs after a property of a shape that is not stored in a cell is changed in a document.|
|[ShapeDataGraphicChanged](application-shapedatagraphicchanged-event-visio.md)|Occurs after a data graphic is applied to or deleted from a shape.|
|[ShapeExitedTextEdit](application-shapeexitedtextedit-event-visio.md)|Occurs after a shape is no longer open for interactive text editing.|
|[ShapeLinkAdded](application-shapelinkadded-event-visio.md)|Occurs after a shape is linked to a data row.|
|[ShapeLinkDeleted](application-shapelinkdeleted-event-visio.md)|Occurs after the link between a shape and a data row is deleted.|
|[ShapeParentChanged](application-shapeparentchanged-event-visio.md)|Occurs after shapes are grouped or a group is ungrouped.|
|[StyleAdded](application-styleadded-event-visio.md)|Occurs after a new style is added to a document.|
|[StyleChanged](application-stylechanged-event-visio.md)|Occurs after the name of a style is changed or a change to the style propagates to objects to which the style is applied.|
|[StyleDeleteCanceled](application-styledeletecanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelStyleDelete** event.|
|[SuspendCanceled](application-suspendcanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelSuspend** event.|
|[SuspendEventsCanceled](application-suspendeventscanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelSuspendEvents** event. .|
|[TextChanged](application-textchanged-event-visio.md)|Occurs after the text of a shape is changed in a document.|
|[UngroupCanceled](application-ungroupcanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelUngroup** event.|
|[ViewChanged](application-viewchanged-event-visio.md)|Occurs when the zoom level or scroll position of a drawing window changes.|
|[VisioIsIdle](application-visioisidle-event-visio.md)|Occurs after the application empties its message queue.|
|[WindowActivated](application-windowactivated-event-visio.md)|Occurs after the active window changes in a Microsoft Visio instance.|
|[WindowChanged](application-windowchanged-event-visio.md)|Occurs when the size or position of a window changes.|
|[WindowCloseCanceled](application-windowclosecanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelWindowClose** event.|
|[WindowOpened](application-windowopened-event-visio.md)|Occurs after a window is opened.|
|[WindowTurnedToPage](application-windowturnedtopage-event-visio.md)|Occurs after a window shows a different page.|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[AddUndoUnit](application-addundounit-method-visio.md)|Adds an object that supports the  **IOleUndoUnit** or **IVBUndoUnit** interface to the Microsoft Visio undo queue.|
|[BeginUndoScope](application-beginundoscope-method-visio.md)|Starts a transaction with a unique scope ID for an instance of Microsoft Visio.|
|[ClearCustomMenus](application-clearcustommenus-method-visio.md)|Restores the built-in Microsoft Visio user interface.|
|[ClearCustomToolbars](application-clearcustomtoolbars-method-visio.md)|Restores the built-in Microsoft Visio user interface.|
|[ConvertResult](application-convertresult-method-visio.md)|Converts a string or number into an equivalent number in different measurement units.|
|[DoCmd](application-docmd-method-visio.md)|Performs the command that has the indicated command ID.|
|[EndUndoScope](application-endundoscope-method-visio.md)|Ends or cancels a transaction that has a unique scope.|
|[EnumDirectories](application-enumdirectories-method-visio.md)|Returns an array naming the folders Microsoft Visio would search, given a list of paths.|
|[FormatResult](application-formatresult-method-visio.md)|Formats a string or number into a string according to a format picture. Uses specified units for scaling and formatting.|
|[FormatResultEx](application-formatresultex-method-visio.md)|Formats a string or number into a string according to a format picture, using specified units for scaling and formatting. Optionally, for date or time strings, sets the language and calendar type of the string.|
|[GetBuiltInStencilFile](application-getbuiltinstencilfile-method-visio.md)|Returns the file path to the specified built-in, hidden stencil used to populate certain galleries in the Microsoft Visio user interface.|
|[GetCustomStencilFile](application-getcustomstencilfile-method-visio.md)|Returns the path to the specified custom stencil used to populate certain galleries in the Microsoft Visio user interface.|
|[GetPreviewEnabled](application-getpreviewenabled-method-visio.md)|Returns a value that indicates whether preview is enabled for the specified gallery in the Microsoft Visio user interface.|
|[InvokeHelp](application-invokehelp-method-visio.md)|Performs operations that involve the Microsoft Visio Help system.|
|[OnComponentEnterState](application-oncomponententerstate-method-visio.md)|Informs a Microsoft Visio instance that client code is causing the instance to enter or exit a particular state.|
|[PurgeUndo](application-purgeundo-method-visio.md)|Empties the Microsoft Visio queue of undo actions.|
|[QueueMarkerEvent](application-queuemarkerevent-method-visio.md)|Queues a  **MarkerEvent** event that fires after all other queued events.|
|[Quit](application-quit-method-visio.md)|Closes the indicated instance of Microsoft Visio.|
|[Redo](application-redo-method-visio.md)|Reverses the most recent undo unit.|
|[RegisterRibbonX](application-registerribbonx-method-visio.md)|Registers the  **[IRibbonExtensibility](iribbonextensibility-object-office.md)** interface that is implemented by the specified add-on to populate the custom user interface (UI).|
|[RenameCurrentScope](application-renamecurrentscope-method-visio.md)|Renames the top-level open undo scope.|
|[SetCustomMenus](application-setcustommenus-method-visio.md)|Replaces the current built-in or custom menus of an application or document.|
|[SetCustomToolbars](application-setcustomtoolbars-method-visio.md)|Replaces the current built-in or custom toolbars of an application or document.|
|[SetPreviewEnabled](application-setpreviewenabled-method-visio.md)|Turns preview on or off for a gallery in the Microsoft Visio user interface.|
|[Undo](application-undo-method-visio.md)|Reverses the most recent undo unit, if the undo unit can be reversed.|
|[UnregisterRibbonX](application-unregisterribbonx-method-visio.md)|Unregisters a previouly registered  **IRibbonExtensiblity** interface that a Microsoft Visio add-in implements.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Active](application-active-property-visio.md)|Indicates whether the instance of Microsoft Visio represented by the  **Application** object is the active application on the Microsoft Windows desktop—the application that has the highlighted title bar. Read-only.|
|[ActiveDocument](application-activedocument-property-visio.md)|Returns the active  **Document** object, which is the document shown in the active window. Read-only.|
|[ActivePage](application-activepage-property-visio.md)|Returns the active  **Page** object. Read-only.|
|[ActivePrinter](application-activeprinter-property-visio.md)|Specifies the printer that all Microsoft Visio documents print to. Read/write.|
|[ActiveWindow](application-activewindow-property-visio.md)|Returns the active  **Window** object. Read-only.|
|[AddonPaths](application-addonpaths-property-visio.md)|Gets or sets the paths where Microsoft Visio looks for third-party or user add-ons. Read/write.|
|[Addons](application-addons-property-visio.md)|Returns the  **Addons** collection of an **Application** or **InvisibleApp** object. Read-only.|
|[AlertResponse](application-alertresponse-property-visio.md)|Determines whether Microsoft Visio shows alerts and modal dialog boxes to the user. Read/write.|
|[Application](application-application-property-visio.md)|Returns the instance of Microsoft Visio that is associated with an object. Read-only.|
|[Assistance](application-assistance-property-visio.md)|Gets a reference to the Microsoft Office (MSO)  **IAssistance** object, which provides a means for developers to create a customized help experience for users within Microsoft Office. Read-only.|
|[AutoLayout](application-autolayout-property-visio.md)|Allows you to temporarily disable automatic layout functionality in Microsoft Visio and then re-enable it after you are finished with an action. Read/write.|
|[AutoRecoverInterval](application-autorecoverinterval-property-visio.md)|Represents the time interval (in minutes) for how often you want to save copies of open documents that have unsaved changes in case of a power failure or an application error. Read/write.|
|[AvailablePrinters](application-availableprinters-property-visio.md)|Returns a list of installed printers. Read-only.|
|[Build](application-build-property-visio.md)|Returns the build number of the running instance. Read-only.|
|[BuiltInMenus](application-builtinmenus-property-visio.md)|Returns a  **UIObject** object that represents a copy of the built-in Microsoft Visio menus and accelerators. Read-only.|
|[BuiltInToolbars](application-builtintoolbars-property-visio.md)|Returns a  **UIObject** object that represents a copy of the built-in Microsoft Visio toolbars. Read-only.|
|[COMAddIns](application-comaddins-property-visio.md)|Returns a reference to the  **COMAddIns** collection that represents all the Component Object Model (COM) add-ins currently registered in Microsoft Visio. Read-only.|
|[CommandBars](application-commandbars-property-visio.md)|Returns a reference to the  **CommandBars** collection that represents the command bars in the container application. Read-only.|
|[CommandLine](application-commandline-property-visio.md)|Determines how Microsoft Visio was started. Read-only.|
|[ConnectorToolDataObject](application-connectortooldataobject-property-visio.md)|Returns an  **IDataObject** interface representing the active **Connector** tool used in the Microsoft Visio user interface. Read-only.|
|[CurrentEdition](application-currentedition-property-visio.md)|Returns an enumerated value that represents the edition of the current instance of Microsoft Visio. Read-only.|
|[CurrentScope](application-currentscope-property-visio.md)|Determines the ID of the scope that causes an event to fire. Read-only.|
|[CustomMenus](application-custommenus-property-visio.md)|Gets a  **UIObject** object that represents the current custom menus and accelerators of an **Application** object. Read-only.|
|[CustomMenusFile](application-custommenusfile-property-visio.md)|Gets or sets the name of the file that defines custom menus and accelerators for an  **Application** object. Read/write.|
|[CustomToolbars](application-customtoolbars-property-visio.md)|Gets a  **UIObject** object that represents the current custom toolbars and status bars of an **Application** object. Read-only.|
|[CustomToolbarsFile](application-customtoolbarsfile-property-visio.md)|Returns or sets the name of the file that defines custom toolbars and status bars for an  **Application** object. Read/write.|
|[DataFeaturesEnabled](application-datafeaturesenabled-property-visio.md)|Gets whether the data features in Microsoft Visio are enabled for the current instance of Visio. Read-only.|
|[DefaultAngleUnits](application-defaultangleunits-property-visio.md)|Determines the default unit of measure for quantities that represent angles. Read/write.|
|[DefaultDurationUnits](application-defaultdurationunits-property-visio.md)|Determines the default unit of measure for quantities that represent durations. Read/write.|
|[DefaultRectangleDataObject](application-defaultrectangledataobject-property-visio.md)|Returns an  **IDataObject** interface that represents the **Rectangle** tool used in the Microsoft Visio user interface. Read-only.|
|[DefaultTextUnits](application-defaulttextunits-property-visio.md)|Determines the default unit of measure for quantities that represent text metrics. Read/write.|
|[DefaultZoomBehavior](application-defaultzoombehavior-property-visio.md)|Determines the zoom behavior for all new Microsoft Visio documents and drawing windows. Read/write.|
|[DeferRecalc](application-deferrecalc-property-visio.md)|Determines whether the application recalculates cell formulas during a series of actions. Read/write.|
|[DeferRelationshipRecalc](application-deferrelationshiprecalc-property-visio.md)|Determines whether Microsoft Visio defers recalculating shape sizes and relationships when a member of the relationship pair is moved or resized. Read/write.|
|[DialogFont](application-dialogfont-property-visio.md)|Returns information about the fonts that Microsoft Visio uses in its dialog boxes. Read-only.|
|[Documents](application-documents-property-visio.md)|Returns the  **Documents** collection for a Microsoft Visio instance. Read-only.|
|[DrawingPaths](application-drawingpaths-property-visio.md)|Gets or sets the paths where Microsoft Visio looks for drawings. Read/write.|
|[EventInfo](application-eventinfo-property-visio.md)|Gets additional information associated with an event, if any exists. Read-only.|
|[EventList](application-eventlist-property-visio.md)|Returns the  **EventList** collection of an object or the **EventList** collection that contains an **Event** object. Read-only.|
|[EventsEnabled](application-eventsenabled-property-visio.md)|Determines whether a Microsoft Visio instance fires events. Read/write.|
|[FullBuild](application-fullbuild-property-visio.md)|Returns the full build number of the running instance. Read-only. |
|[HelpPaths](application-helppaths-property-visio.md)|Gets or sets the paths where Microsoft Visio looks for Help files. Read/write.|
|[InhibitSelectChange](application-inhibitselectchange-property-visio.md)|Determines whether shapes added to the drawing page by Automation are selected. Read/write.|
|[InstanceHandle32](application-instancehandle32-property-visio.md)|Gets the instance handle of the  **Application** object for a 32-bit version of Microsoft Visio. Read-only.|
|[InstanceHandle64](application-instancehandle64-property-visio.md)|Gets the instance handle of the  **[Application](application-object-visio.md)** object for a 64-bit version of Microsoft Visio. Read-only.|
|[IsInScope](application-isinscope-property-visio.md)|Determines whether a call to an event handler is between an  **EnterScope** event and an **ExitScope** event for a scope. Read-only.|
|[IsUndoingOrRedoing](application-isundoingorredoing-property-visio.md)|Determines whether the current event handler is being called as a result of an  **Undo** or **Redo** action in the application. Read-only.|
|[IsVisio32](application-isvisio32-property-visio.md)|Returns  **True** if the current instance of Microsoft Visio is 32-bit. Returns **False** if the current instance is 64-bit. Read-only.|
|[Language](application-language-property-visio.md)|Represents the language ID of the version of the Microsoft Visio instance represented by the parent object. Read/write.|
|[LanguageHelp](application-languagehelp-property-visio.md)|Represents the language ID of the Help in the version of the Microsoft Visio instance represented by the parent object. Read-only.|
|[LanguageSettings](application-languagesettings-property-visio.md)|Returns a reference to the Microsoft Office (MSO)  **LanguageSettings** interface. Read-only.|
|[LiveDynamics](application-livedynamics-property-visio.md)|Controls whether Microsoft Visio recalculates shape properties during drag operations on every mouse move or only after the mouse button is released. Read/write.|
|[MyShapesPath](application-myshapespath-property-visio.md)|Gets or sets where Microsoft Visio looks for the  **My Shapes** folder on the user's hard disk. Read/write.|
|[Name](application-name-property-visio.md)|Specifies the name of an object. Read-only.|
|[ObjectType](application-objecttype-property-visio.md)|Returns an object's type. Read-only.|
|[OnDataChangeDelay](application-ondatachangedelay-property-visio.md)|Gets or sets how long the Microsoft Visio instance waits before advising a container application that a Visio document being shown by the container has changed and should be redisplayed. Read/write.|
|[Path](application-path-property-visio.md)|Returns the drive and folder path of the Microsoft Visio application. Read-only.|
|[PersistsEvents](application-persistsevents-property-visio.md)|Indicates whether an object is capable of containing persistent events in its  **EventList** collection. Read-only.|
|[ProcessID](application-processid-property-visio.md)|Returns the unique identity of the current Microsoft Visio process. Read-only.|
|[PromptForSummary](application-promptforsummary-property-visio.md)|Determines whether Microsoft Visio prompts for document properties when it saves a document. Read/write.|
|[SaveAsWebObject](application-saveaswebobject-property-visio.md)|Returns a reference to the  **IDispatch** interface of a **VisSaveAsWeb** object. Read-only.|
|[ScreenUpdating](application-screenupdating-property-visio.md)|Determines whether the screen is updated (redrawn) during a series of actions. Read/write.|
|[Settings](application-settings-property-visio.md)|Returns an  **ApplicationSettings** object, which you can use to set Microsoft Visio application properties. Read-only.|
|[ShowChanges](application-showchanges-property-visio.md)|Determines whether the screen is updated (redrawn) during a series of actions. Read/write.|
|[ShowProgress](application-showprogress-property-visio.md)|Determines whether a progress indicator is shown while performing certain operations. Read/write.|
|[ShowStatusBar](application-showstatusbar-property-visio.md)|Determines whether the status bar is shown. Read/write.|
|[ShowToolbar](application-showtoolbar-property-visio.md)|Determines whether toolbars and menu bars are visible. Read/write.|
|[StartupPaths](application-startuppaths-property-visio.md)|Gets or sets the paths where Microsoft Visio looks for third-party and user add-ons to run when the application is started. Read/write.|
|[Stat](application-stat-property-visio.md)|Returns status information for an object. Read-only.|
|[StencilPaths](application-stencilpaths-property-visio.md)|Gets or sets the paths where Microsoft Visio looks for stencils. Read/write.|
|[TemplatePaths](application-templatepaths-property-visio.md)|Gets or sets the paths where Microsoft Visio looks for templates. Read/write.|
|[TraceFlags](application-traceflags-property-visio.md)|Gets or sets events logged during a Microsoft Visio instance. Read/write.|
|[TypelibMajorVersion](application-typelibmajorversion-property-visio.md)|Returns the major version number of the Microsoft Visio type library. Read-only.|
|[TypelibMinorVersion](application-typelibminorversion-property-visio.md)|Returns the minor version number of the Microsoft Visio type library. Read-only.|
|[UndoEnabled](application-undoenabled-property-visio.md)|Determines whether undo information is maintained in memory. Read/write.|
|[UserName](application-username-property-visio.md)|Gets or sets the user name of an  **Application** object. Read/write.|
|[VBAEnabled](application-vbaenabled-property-visio.md)|Specifies whether Microsoft Visual Basic for Applications (VBA) is enabled in the application. Read-only.|
|[Vbe](application-vbe-property-visio.md)|Gets the root object of the object model exposed by Microsoft Visual Basic for Applications (VBA). Use this property to access and manipulate the VBA projects associated with currently open Microsoft Visio documents. Read-only.|
|[Version](application-version-property-visio.md)|Returns the version of a running Microsoft Visio instance. Read-only.|
|[Visible](application-visible-property-visio.md)|Determines whether an object is visible. Read/write.|
|[Window](application-window-property-visio.md)|Returns the window associated with the current instance of Microsoft Visio. Read-only.|
|[WindowHandle32](application-windowhandle32-property-visio.md)|Returns the 32-bit handle of a Microsoft Visio window. Read-only.|
|[Windows](application-windows-property-visio.md)|Returns the  **Windows** collection for a Microsoft Visio instance or window. Read-only.|

