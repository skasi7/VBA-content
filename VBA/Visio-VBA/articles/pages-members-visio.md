---
title: Pages Members (Visio)
ms.prod: VISIO
ms.assetid: 49e81797-27d2-2005-efb0-e00f973fbbc9
---


# Pages Members (Visio)
Includes a  **Page** object for each drawing page in a document.

Includes a  **Page** object for each drawing page in a document.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[AfterReplaceShapes](pages-afterreplaceshapes-event-visio.md)|Occurs after a shape-replacement operation.|
|[BeforePageDelete](pages-beforepagedelete-event-visio.md)|Occurs before a page is deleted.|
|[BeforeReplaceShapes](pages-beforereplaceshapes-event-visio.md)|Occurs just before a shape-replacement operation.|
|[BeforeSelectionDelete](pages-beforeselectiondelete-event-visio.md)|Occurs before selected objects are deleted.|
|[BeforeShapeDelete](pages-beforeshapedelete-event-visio.md)|Occurs before a shape is deleted.|
|[BeforeShapeTextEdit](pages-beforeshapetextedit-event-visio.md)|Occurs before a shape is opened for text editing in the user interface.|
|[CalloutRelationshipAdded](pages-calloutrelationshipadded-event-visio.md)|Occurs when a new callout relationship is added to a page.|
|[CalloutRelationshipDeleted](pages-calloutrelationshipdeleted-event-visio.md)|Occurs when a callout relationship is deleted from a page.|
|[CellChanged](pages-cellchanged-event-visio.md)|Occurs after the value changes in a cell in a document.|
|[ConnectionsAdded](pages-connectionsadded-event-visio.md)|Occurs after connections have been established between shapes.|
|[ConnectionsDeleted](pages-connectionsdeleted-event-visio.md)|Occurs after connections between shapes have been removed.|
|[ContainerRelationshipAdded](pages-containerrelationshipadded-event-visio.md)|Occurs when a new container relationship is added to the document.|
|[ContainerRelationshipDeleted](pages-containerrelationshipdeleted-event-visio.md)|Occurs when a container relationship is deleted from the document.|
|[ConvertToGroupCanceled](pages-converttogroupcanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelConvertToGroup** event.|
|[FormulaChanged](pages-formulachanged-event-visio.md)|Occurs after a formula changes in a cell in the object that receives the event.|
|[GroupCanceled](pages-groupcanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelGroup** event.|
|[PageAdded](pages-pageadded-event-visio.md)|Occurs after a new page is added to a document.|
|[PageChanged](pages-pagechanged-event-visio.md)|Occurs after the name of a page, the background page associated with a page, or the page type (foreground or background) changes.|
|[PageDeleteCanceled](pages-pagedeletecanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelPageDelete** event.|
|[QueryCancelConvertToGroup](pages-querycancelconverttogroup-event-visio.md)|Occurs before the application converts a selection of shapes to a group in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelGroup](pages-querycancelgroup-event-visio.md)|Occurs before the application groups a selection of shapes in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelPageDelete](pages-querycancelpagedelete-event-visio.md)|Occurs before the application deletes a page in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelReplaceShapes](pages-querycancelreplaceshapes-event-visio.md)|Occurs immediately after a shape-replacement operation is requested. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelSelectionDelete](pages-querycancelselectiondelete-event-visio.md)|Occurs before the application deletes a selection of shapes in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelUngroup](pages-querycancelungroup-event-visio.md)|Occurs before the application ungroups a selection of shapes in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[ReplaceShapesCanceled](pages-replaceshapescanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelReplaceShapes** event.|
|[SelectionAdded](pages-selectionadded-event-visio.md)|Occurs after one or more shapes are added to a document.|
|[SelectionDeleteCanceled](pages-selectiondeletecanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelSelectionDelete** event.|
|[ShapeAdded](pages-shapeadded-event-visio.md)|Occurs after one or more shapes are added to a document.|
|[ShapeChanged](pages-shapechanged-event-visio.md)|Occurs after a property of a shape that is not stored in a cell is changed in a document.|
|[ShapeDataGraphicChanged](pages-shapedatagraphicchanged-event-visio.md)|Occurs after a data graphic is applied to or deleted from a shape.|
|[ShapeExitedTextEdit](pages-shapeexitedtextedit-event-visio.md)|Occurs after a shape is no longer open for interactive text editing.|
|[ShapeLinkAdded](pages-shapelinkadded-event-visio.md)|Occurs after a shape is linked to a data row.|
|[ShapeLinkDeleted](pages-shapelinkdeleted-event-visio.md)|Occurs after the link between a shape and a data row is deleted.|
|[ShapeParentChanged](pages-shapeparentchanged-event-visio.md)|Occurs after shapes are grouped or a group is ungrouped.|
|[TextChanged](pages-textchanged-event-visio.md)|Occurs after the text of a shape is changed in a document.|
|[UngroupCanceled](pages-ungroupcanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelUngroup** event.|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Add](pages-add-method-visio.md)|Adds a new object to a collection.|
|[GetNames](pages-getnames-method-visio.md)|Returns the names of all items in a collection.|
|[GetNamesU](pages-getnamesu-method-visio.md)|Returns the universal names of all items in a collection.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](pages-application-property-visio.md)|Returns the instance of Microsoft Visio that is associated with an object. Read-only.|
|[Count](pages-count-property-visio.md)|Returns the number of objects in a collection. Read-only.|
|[Document](pages-document-property-visio.md)|Gets the  **Document** object that is associated with an object. Read-only.|
|[EventList](pages-eventlist-property-visio.md)|Returns the  **EventList** collection of an object or the **EventList** collection that contains an **Event** object. Read-only.|
|[Item](pages-item-property-visio.md)|Returns an item from a collection. The  **Item** property is the default property for all collections. Read-only.|
|[ItemFromID](pages-itemfromid-property-visio.md)|Returns an item of a collection using the ID of the item. Read-only.|
|[ItemU](pages-itemu-property-visio.md)|Returns an object from a collection. Read-only.|
|[ObjectType](pages-objecttype-property-visio.md)|Returns an object's type. Read-only.|
|[PersistsEvents](pages-persistsevents-property-visio.md)|Indicates whether an object is capable of containing persistent events in its  **EventList** collection. Read-only.|
|[Stat](pages-stat-property-visio.md)|Returns status information for an object. Read-only.|

