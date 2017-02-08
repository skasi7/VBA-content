---
title: Window Members (Visio)
ms.prod: VISIO
ms.assetid: 048ce9a3-3650-c4a2-2dc4-e7538ae18d34
---


# Window Members (Visio)
Represents an open window in a Microsoft Visio instance.

Represents an open window in a Microsoft Visio instance.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[BeforeWindowClosed](window-beforewindowclosed-event-visio.md)|Occurs before a window is closed.|
|[BeforeWindowPageTurn](window-beforewindowpageturn-event-visio.md)|Occurs before a window is about to show a different page.|
|[BeforeWindowSelDelete](window-beforewindowseldelete-event-visio.md)|Occurs before the shapes in the selection of a window are deleted.|
|[KeyDown](window-keydown-event-visio.md)|Occurs when a keyboard key is pressed.|
|[KeyPress](window-keypress-event-visio.md)|Occurs when a keyboard key is pressed.|
|[KeyUp](window-keyup-event-visio.md)|Occurs when a keyboard key is released.|
|[MouseDown](window-mousedown-event-visio.md)|Occurs when a mouse button is clicked.|
|[MouseMove](window-mousemove-event-visio.md)|Occurs when the mouse is moved.|
|[MouseUp](window-mouseup-event-visio.md)|Occurs when a mouse button is released.|
|[OnKeystrokeMessageForAddon](window-onkeystrokemessageforaddon-event-visio.md)|Occurs when Microsoft Visio receives a keystroke message from Microsoft Windows that is targeted at an add-on window or child of an add-on window.|
|[QueryCancelWindowClose](window-querycancelwindowclose-event-visio.md)|Occurs before the application closes a window in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[SelectionChanged](window-selectionchanged-event-visio.md)|Occurs after a set of shapes selected in a window changes.|
|[ViewChanged](window-viewchanged-event-visio.md)|Occurs when the zoom level or scroll position of a drawing window changes.|
|[WindowActivated](window-windowactivated-event-visio.md)|Occurs after the active window changes in a Microsoft Visio instance.|
|[WindowChanged](window-windowchanged-event-visio.md)|Occurs when the size or position of a window changes.|
|[WindowCloseCanceled](window-windowclosecanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelWindowClose** event.|
|[WindowTurnedToPage](window-windowturnedtopage-event-visio.md)|Occurs after a window shows a different page.|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Activate](window-activate-method-visio.md)|Activates a window.|
|[CenterViewOnShape](window-centerviewonshape-method-visio.md)|Pans the Microsoft Visio drawing window to place the specified shape in the center of the view.|
|[Close](window-close-method-visio.md)|Closes a window.|
|[DeselectAll](window-deselectall-method-visio.md)|Deselects all shapes in a window or selection.|
|[DockedStencils](window-dockedstencils-method-visio.md)|Returns the names of all stencils docked in a Microsoft Visio drawing window.|
|[GetViewRect](window-getviewrect-method-visio.md)|Returns the page coordinates of a window's borders.|
|[GetWindowRect](window-getwindowrect-method-visio.md)|Gets the size and position of the client area of a window.|
|[NewWindow](window-newwindow-method-visio.md)|Opens a new Microsoft Visio window.|
|[Scroll](window-scroll-method-visio.md)|Scrolls the contents of a window vertically, horizontally, or both.|
|[ScrollViewTo](window-scrollviewto-method-visio.md)|Scrolls a window to a particular page coordinate.|
|[Select](window-select-method-visio.md)|Selects or clears an object.|
|[SelectAll](window-selectall-method-visio.md)|Selects all possible shapes in a window or selection.|
|[SetViewRect](window-setviewrect-method-visio.md)|Sets the page coordinates of a window's borders by adjusting the zoom level and center scroll position.|
|[SetWindowRect](window-setwindowrect-method-visio.md)|Sets the size and position of the client area of a window.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[AllowEditing](window-allowediting-property-visio.md)|Determines whether the  **Edit Stencil** command is enabled or disabled in a stencil window. Read/write.|
|[Application](window-application-property-visio.md)|Returns the instance of Microsoft Visio that is associated with an object. Read-only.|
|[BackgroundColor](window-backgroundcolor-property-visio.md)|Determines the background color of the active Microsoft Visio drawing window and its associated print preview and full-screen view windows. Read/write.|
|[BackgroundColorGradient](window-backgroundcolorgradient-property-visio.md)|Determines the background gradient color of the active Microsoft Visio drawing window and its associated print preview and full-screen view windows. Read/write.|
|[Caption](window-caption-property-visio.md)|Gets or sets the caption for a window. Read/write.|
|[Document](window-document-property-visio.md)|Gets the  **Document** object that is associated with an object. Read-only.|
|[EventList](window-eventlist-property-visio.md)|Returns the  **EventList** collection of an object or the **EventList** collection that contains an **Event** object. Read-only.|
|[ID](window-id-property-visio.md)|Gets the ID of an object. Read-only.|
|[Index](window-index-property-visio.md)|Gets the ordinal position of a  **Window** object in the **Windows** collection. Read-only.|
|[InPlace](window-inplace-property-visio.md)|Specifies whether a window is open in place, or whether a document is being viewed through a window that is open in place. Read-only.|
|[IsEditingOLE](window-iseditingole-property-visio.md)|Determines whether a drawing window contains an ActiveX control that has focus or an embedded or linked object that is being edited. Read-only.|
|[IsEditingText](window-iseditingtext-property-visio.md)|Determines whether a text editing session is active in the drawing window. Read-only.|
|[Master](window-master-property-visio.md)|Gets the master that is displayed in a window. Read-only.|
|[MasterShortcut](window-mastershortcut-property-visio.md)|Gets the master shortcut that is displayed in a window. Read-only.|
|[MergeCaption](window-mergecaption-property-visio.md)|Gets or sets the abbreviated caption that appears on the page tab when the window is merged with other windows. Read/write.|
|[MergeClass](window-mergeclass-property-visio.md)|Specifies a list of window classes that this anchored window can merge with. Read/write.|
|[MergeID](window-mergeid-property-visio.md)|Specifies the string version of a merged window's globally unique identifier (GUID). Read/write.|
|[MergePosition](window-mergeposition-property-visio.md)|Specifies the left-to-right tab position of a merged anchored window. Read/write.|
|[ObjectType](window-objecttype-property-visio.md)|Returns an object's type. Read-only.|
|[Page](window-page-property-visio.md)|Gets or sets the page that is displayed in a window. Read/write.|
|[Parent](window-parent-property-visio.md)|Determines the parent of an object. Read-only.|
|[ParentWindow](window-parentwindow-property-visio.md)|Returns the  **Window** object that is the parent of another **Window** object. Read-only.|
|[PersistsEvents](window-persistsevents-property-visio.md)|Indicates whether an object is capable of containing persistent events in its  **EventList** collection. Read-only.|
|[ReviewerMarkupVisible](window-reviewermarkupvisible-property-visio.md)|Determines whether reviewer markup, for a particular reviewer or all reviewers, is visible in a Microsoft Visio window that displays a drawing page. Read/write.|
|[ScrollLock](window-scrolllock-property-visio.md)|Determines whether scrolling is disabled in a Microsoft Visio window. Read/write.|
|[SelectedCell](window-selectedcell-property-visio.md)|Returns the selected cell in the ShapeSheet window. Read-only.|
|[SelectedDataRecordset](window-selecteddatarecordset-property-visio.md)|Gets or sets the data recordset that is displayed on the active tab of the  **External Data Window** in the Microsoft Visio user interface (UI). Read/write.|
|[SelectedDataRowID](window-selecteddatarowid-property-visio.md)|Gets or sets the ID of the data row that is selected (or that is the primary row selected, when multiple rows are selected) on the active tab of the  **External Data Window** in the Microsoft Visio user interface (UI). Read/write.|
|[SelectedMasters](window-selectedmasters-property-visio.md)| Returns an array of the masters or master shortcuts selected in a Microsoft Visio stencil window. Read-only.|
|[SelectedText](window-selectedtext-property-visio.md)|Returns the selected text in the Microsoft Visio drawing window as a  **Characters** object. Read/write.|
|[SelectedValidationIssue](window-selectedvalidationissue-property-visio.md)|Gets or sets the validation issue that is selected in the  **Issues** window. Read/write.|
|[Selection](window-selection-property-visio.md)|Returns a  **Selection** object that represents what is presently selected in the window, or assigns a selection created by the **CreateSelection** method to a **Selection** object. Read/write.|
|[SelectionForDragCopy](window-selectionfordragcopy-property-visio.md)|Returns the  **[Selection](selection-object-visio.md)** object that represents the collection of shapes that will participate in drag or copy operations, based on the current selection. Read-only.|
|[Shape](window-shape-property-visio.md)|Returns the  **Shape** object that owns a **Cell** , **Characters** , **Row** , or **Section** object or that is associated with a **Hyperlink** or **OLEObject** object or with the **Hyperlinks** collection. Read-only.|
|[ShowConnectPoints](window-showconnectpoints-property-visio.md)|Determines whether connection points are shown in a window. Read/write.|
|[ShowGrid](window-showgrid-property-visio.md)|Determines whether a grid is shown in a window. Read/write.|
|[ShowGuides](window-showguides-property-visio.md)|Determines whether guides are shown in a window. Read/write.|
|[ShowPageBreaks](window-showpagebreaks-property-visio.md)|Determines whether page breaks are shown in a window. Read/write.|
|[ShowPageOutline](window-showpageoutline-property-visio.md)|Determines whether the drawing page outline is displayed in the Microsoft Visio drawing window. Read/write.|
|[ShowPageTabs](window-showpagetabs-property-visio.md)|Determines whether page tab controls are shown in the drawing window. Read/write.|
|[ShowRulers](window-showrulers-property-visio.md)|Determines whether rulers are shown in the drawing window. Read/write.|
|[ShowScrollBars](window-showscrollbars-property-visio.md)|Determines whether scroll bars are shown in the drawing window. Read/write.|
|[Stat](window-stat-property-visio.md)|Returns status information for an object. Read-only.|
|[SubType](window-subtype-property-visio.md)|Returns the subtype of a  **Window** object that represents a drawing window. Read-only.|
|[Type](window-type-property-visio.md)|Returns the type of the object. Read-only.|
|[ViewFit](window-viewfit-property-visio.md)|Determines which auto-fit mode a window is in, if any. Read/write.|
|[Visible](window-visible-property-visio.md)|Determines whether a window is visible. Read/write.|
|[WindowHandle32](window-windowhandle32-property-visio.md)|Returns the 32-bit handle of a Microsoft Visio window. Read-only.|
|[Windows](window-windows-property-visio.md)|Returns the  **Windows** collection for a Microsoft Visio instance or window. Read-only.|
|[WindowState](window-windowstate-property-visio.md)|Gets or sets the state of a window. Read/write.|
|[Zoom](window-zoom-property-visio.md)|Gets or sets the current display size (magnification factor) for a page in a window. Read/write.|
|[ZoomBehavior](window-zoombehavior-property-visio.md)|Determines the zoom behavior for a Microsoft Visio document or window. Read/write.|
|[ZoomLock](window-zoomlock-property-visio.md)|Determines whether zooming is disabled in a Microsoft Visio drawing window. Read/write.|

