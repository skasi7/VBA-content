---
title: Masters Members (Visio)
ms.prod: VISIO
ms.assetid: 3f1db5a1-0b7b-c0ac-c9af-1e7ac97d5e64
---


# Masters Members (Visio)
 Includes a **Master** object for each master in a document's stencil.

 Includes a **Master** object for each master in a document's stencil.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[BeforeMasterDelete](masters-beforemasterdelete-event-visio.md)|Occurs before a master is deleted from a document.|
|[BeforeSelectionDelete](masters-beforeselectiondelete-event-visio.md)|Occurs before selected objects are deleted.|
|[BeforeShapeDelete](masters-beforeshapedelete-event-visio.md)|Occurs before a shape is deleted.|
|[BeforeShapeTextEdit](masters-beforeshapetextedit-event-visio.md)|Occurs before a shape is opened for text editing in the user interface.|
|[CellChanged](masters-cellchanged-event-visio.md)|Occurs after the value changes in a cell in a document.|
|[ConnectionsAdded](masters-connectionsadded-event-visio.md)|Occurs after connections have been established between shapes.|
|[ConnectionsDeleted](masters-connectionsdeleted-event-visio.md)|Occurs after connections between shapes have been removed.|
|[ConvertToGroupCanceled](masters-converttogroupcanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelConvertToGroup** event.|
|[FormulaChanged](masters-formulachanged-event-visio.md)|Occurs after a formula changes in a cell in the object that receives the event.|
|[GroupCanceled](masters-groupcanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelGroup** event.|
|[MasterAdded](masters-masteradded-event-visio.md)|Occurs after a new master is added to a document.|
|[MasterChanged](masters-masterchanged-event-visio.md)|Occurs after properties of a master are changed and propagated to its instances.|
|[MasterDeleteCanceled](masters-masterdeletecanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelMasterDelete** event.|
|[QueryCancelConvertToGroup](masters-querycancelconverttogroup-event-visio.md)|Occurs before the application converts a selection of shapes to a group in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelGroup](masters-querycancelgroup-event-visio.md)|Occurs before the application groups a selection of shapes in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelMasterDelete](masters-querycancelmasterdelete-event-visio.md)|Occurs before the application deletes a master in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelSelectionDelete](masters-querycancelselectiondelete-event-visio.md)|Occurs before the application deletes a selection of shapes in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelUngroup](masters-querycancelungroup-event-visio.md)|Occurs before the application ungroups a selection of shapes in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[SelectionAdded](masters-selectionadded-event-visio.md)|Occurs after one or more shapes are added to a document.|
|[SelectionDeleteCanceled](masters-selectiondeletecanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelSelectionDelete** event.|
|[ShapeAdded](masters-shapeadded-event-visio.md)|Occurs after one or more shapes are added to a document.|
|[ShapeChanged](masters-shapechanged-event-visio.md)|Occurs after a property of a shape that is not stored in a cell is changed in a document.|
|[ShapeDataGraphicChanged](masters-shapedatagraphicchanged-event-visio.md)|Occurs after a data graphic is applied to or deleted from a shape.|
|[ShapeExitedTextEdit](masters-shapeexitedtextedit-event-visio.md)|Occurs after a shape is no longer open for interactive text editing.|
|[ShapeParentChanged](masters-shapeparentchanged-event-visio.md)|Occurs after shapes are grouped or a group is ungrouped.|
|[TextChanged](masters-textchanged-event-visio.md)|Occurs after the text of a shape is changed in a document.|
|[UngroupCanceled](masters-ungroupcanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelUngroup** event.|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Add](masters-add-method-visio.md)|Adds a new object to a collection.|
|[AddEx](masters-addex-method-visio.md)|Adds a new  **Master** object of the specified type to the **Masters** collection of a Microsoft Visio document.|
|[Drop](masters-drop-method-visio.md)|Creates a new **Master** object by dropping an object onto a receiving object such as a stencil or document, or the **Masters** or **MasterShortcuts** collection.|
|[GetNames](masters-getnames-method-visio.md)|Returns the names of all items in a collection.|
|[GetNamesU](masters-getnamesu-method-visio.md)|Returns the universal names of all items in a collection.|
|[Paste](masters-paste-method-visio.md)|Pastes the contents of the Clipboard into an object.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](masters-application-property-visio.md)|Returns the instance of Microsoft Visio that is associated with an object. Read-only.|
|[Count](masters-count-property-visio.md)|Returns the number of objects in a collection. Read-only.|
|[Document](masters-document-property-visio.md)|Gets the  **Document** object that is associated with an object. Read-only.|
|[EventList](masters-eventlist-property-visio.md)|Returns the  **EventList** collection of an object or the **EventList** collection that contains an **Event** object. Read-only.|
|[Item](masters-item-property-visio.md)|Returns an item from a collection. The  **Item** property is the default property for all collections. Read-only.|
|[ItemFromID](masters-itemfromid-property-visio.md)|Returns an item of a collection using the ID of the item. Read-only.|
|[ItemU](masters-itemu-property-visio.md)|Returns an object from a collection. Read-only.|
|[ObjectType](masters-objecttype-property-visio.md)|Returns an object's type. Read-only.|
|[PersistsEvents](masters-persistsevents-property-visio.md)|Indicates whether an object is capable of containing persistent events in its  **EventList** collection. Read-only.|
|[Stat](masters-stat-property-visio.md)|Returns status information for an object. Read-only.|

