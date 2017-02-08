---
title: Windows Members (Visio)
ms.prod: VISIO
ms.assetid: e2f07639-408d-346f-85b6-ebf686239b40
---


# Windows Members (Visio)
 Includes a **Window** object for a window that is open in the application.

 Includes a **Window** object for a window that is open in the application.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[BeforeWindowClosed](windows-beforewindowclosed-event-visio.md)|Occurs before a window is closed.|
|[BeforeWindowPageTurn](windows-beforewindowpageturn-event-visio.md)|Occurs before a window is about to show a different page.|
|[BeforeWindowSelDelete](windows-beforewindowseldelete-event-visio.md)|Occurs before the shapes in the selection of a window are deleted.|
|[KeyDown](windows-keydown-event-visio.md)|Occurs when a keyboard key is pressed.|
|[KeyPress](windows-keypress-event-visio.md)|Occurs when a keyboard key is pressed.|
|[KeyUp](windows-keyup-event-visio.md)|Occurs when a keyboard key is released.|
|[MouseDown](windows-mousedown-event-visio.md)|Occurs when a mouse button is clicked.|
|[MouseMove](windows-mousemove-event-visio.md)|Occurs when the mouse is moved.|
|[MouseUp](windows-mouseup-event-visio.md)|Occurs when a mouse button is released.|
|[OnKeystrokeMessageForAddon](windows-onkeystrokemessageforaddon-event-visio.md)|Occurs when Microsoft Visio receives a keystroke message from Microsoft Windows that is targeted at an add-on window or child of an add-on window.|
|[QueryCancelWindowClose](windows-querycancelwindowclose-event-visio.md)|Occurs before the application closes a window in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[SelectionChanged](windows-selectionchanged-event-visio.md)|Occurs after a set of shapes selected in a window changes.|
|[ViewChanged](windows-viewchanged-event-visio.md)|Occurs when the zoom level or scroll position of a drawing window changes.|
|[WindowActivated](windows-windowactivated-event-visio.md)|Occurs after the active window changes in a Microsoft Visio instance.|
|[WindowChanged](windows-windowchanged-event-visio.md)|Occurs when the size or position of a window changes.|
|[WindowCloseCanceled](windows-windowclosecanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelWindowClose** event.|
|[WindowOpened](windows-windowopened-event-visio.md)|Occurs after a window is opened.|
|[WindowTurnedToPage](windows-windowturnedtopage-event-visio.md)|Occurs after a window shows a different page.|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Add](windows-add-method-visio.md)|Adds a new  **Window** object to the **Windows** collection.|
|[Arrange](windows-arrange-method-visio.md)|Arranges the windows in a  **Windows** collection.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](windows-application-property-visio.md)|Returns the instance of Microsoft Visio that is associated with an object. Read-only.|
|[Count](windows-count-property-visio.md)|Returns the number of objects in a collection. Read-only.|
|[EventList](windows-eventlist-property-visio.md)|Returns the  **EventList** collection of an object or the **EventList** collection that contains an **Event** object. Read-only.|
|[Item](windows-item-property-visio.md)|Returns an item from a collection. The  **Item** property is the default property for all collections. Read-only.|
|[ItemEx](windows-itemex-property-visio.md)|Returns a  **Window** object from a collection. Read-only.|
|[ItemFromID](windows-itemfromid-property-visio.md)|Returns an item of a collection using the ID of the item. Read-only.|
|[ObjectType](windows-objecttype-property-visio.md)|Returns an object's type. Read-only.|
|[PersistsEvents](windows-persistsevents-property-visio.md)|Indicates whether an object is capable of containing persistent events in its  **EventList** collection. Read-only.|

