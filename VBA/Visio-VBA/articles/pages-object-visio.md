---
title: Pages Object (Visio)
keywords: vis_sdr.chm10195
f1_keywords:
- vis_sdr.chm10195
ms.prod: VISIO
api_name:
- Visio.Pages
ms.assetid: 45eec568-b5cc-5e80-ff5c-4dfa567efb5d
---


# Pages Object (Visio)

Includes a  **Page** object for each drawing page in a document.


## Remarks

To retrieve a  **Pages** collection, use the **Pages** property of a **Document** object.

The default property of a  **Pages** collection is **Item**.

The order of items in a  **Pages** collection is significant: if there are _n_ foreground pages in a document, the first _n_ pages in its **Pages** collection are foreground pages and are in order. The remaining pages in the collection are the background pages of the document; these are in no particular order.

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this collection maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVPages**
    

## Events



|**Name**|
|:-----|
|[AfterReplaceShapes](http://msdn.microsoft.com/library/pages-afterreplaceshapes-event-visio%28Office.15%29.aspx)|
|[BeforePageDelete](http://msdn.microsoft.com/library/pages-beforepagedelete-event-visio%28Office.15%29.aspx)|
|[BeforeReplaceShapes](http://msdn.microsoft.com/library/pages-beforereplaceshapes-event-visio%28Office.15%29.aspx)|
|[BeforeSelectionDelete](http://msdn.microsoft.com/library/pages-beforeselectiondelete-event-visio%28Office.15%29.aspx)|
|[BeforeShapeDelete](http://msdn.microsoft.com/library/pages-beforeshapedelete-event-visio%28Office.15%29.aspx)|
|[BeforeShapeTextEdit](http://msdn.microsoft.com/library/pages-beforeshapetextedit-event-visio%28Office.15%29.aspx)|
|[CalloutRelationshipAdded](http://msdn.microsoft.com/library/pages-calloutrelationshipadded-event-visio%28Office.15%29.aspx)|
|[CalloutRelationshipDeleted](http://msdn.microsoft.com/library/pages-calloutrelationshipdeleted-event-visio%28Office.15%29.aspx)|
|[CellChanged](http://msdn.microsoft.com/library/pages-cellchanged-event-visio%28Office.15%29.aspx)|
|[ConnectionsAdded](http://msdn.microsoft.com/library/pages-connectionsadded-event-visio%28Office.15%29.aspx)|
|[ConnectionsDeleted](http://msdn.microsoft.com/library/pages-connectionsdeleted-event-visio%28Office.15%29.aspx)|
|[ContainerRelationshipAdded](http://msdn.microsoft.com/library/pages-containerrelationshipadded-event-visio%28Office.15%29.aspx)|
|[ContainerRelationshipDeleted](http://msdn.microsoft.com/library/pages-containerrelationshipdeleted-event-visio%28Office.15%29.aspx)|
|[ConvertToGroupCanceled](http://msdn.microsoft.com/library/pages-converttogroupcanceled-event-visio%28Office.15%29.aspx)|
|[FormulaChanged](http://msdn.microsoft.com/library/pages-formulachanged-event-visio%28Office.15%29.aspx)|
|[GroupCanceled](http://msdn.microsoft.com/library/pages-groupcanceled-event-visio%28Office.15%29.aspx)|
|[PageAdded](http://msdn.microsoft.com/library/pages-pageadded-event-visio%28Office.15%29.aspx)|
|[PageChanged](http://msdn.microsoft.com/library/pages-pagechanged-event-visio%28Office.15%29.aspx)|
|[PageDeleteCanceled](http://msdn.microsoft.com/library/pages-pagedeletecanceled-event-visio%28Office.15%29.aspx)|
|[QueryCancelConvertToGroup](http://msdn.microsoft.com/library/pages-querycancelconverttogroup-event-visio%28Office.15%29.aspx)|
|[QueryCancelGroup](http://msdn.microsoft.com/library/pages-querycancelgroup-event-visio%28Office.15%29.aspx)|
|[QueryCancelPageDelete](http://msdn.microsoft.com/library/pages-querycancelpagedelete-event-visio%28Office.15%29.aspx)|
|[QueryCancelReplaceShapes](http://msdn.microsoft.com/library/pages-querycancelreplaceshapes-event-visio%28Office.15%29.aspx)|
|[QueryCancelSelectionDelete](http://msdn.microsoft.com/library/pages-querycancelselectiondelete-event-visio%28Office.15%29.aspx)|
|[QueryCancelUngroup](http://msdn.microsoft.com/library/pages-querycancelungroup-event-visio%28Office.15%29.aspx)|
|[ReplaceShapesCanceled](http://msdn.microsoft.com/library/pages-replaceshapescanceled-event-visio%28Office.15%29.aspx)|
|[SelectionAdded](http://msdn.microsoft.com/library/pages-selectionadded-event-visio%28Office.15%29.aspx)|
|[SelectionDeleteCanceled](http://msdn.microsoft.com/library/pages-selectiondeletecanceled-event-visio%28Office.15%29.aspx)|
|[ShapeAdded](http://msdn.microsoft.com/library/pages-shapeadded-event-visio%28Office.15%29.aspx)|
|[ShapeChanged](http://msdn.microsoft.com/library/pages-shapechanged-event-visio%28Office.15%29.aspx)|
|[ShapeDataGraphicChanged](http://msdn.microsoft.com/library/pages-shapedatagraphicchanged-event-visio%28Office.15%29.aspx)|
|[ShapeExitedTextEdit](http://msdn.microsoft.com/library/pages-shapeexitedtextedit-event-visio%28Office.15%29.aspx)|
|[ShapeLinkAdded](http://msdn.microsoft.com/library/pages-shapelinkadded-event-visio%28Office.15%29.aspx)|
|[ShapeLinkDeleted](http://msdn.microsoft.com/library/pages-shapelinkdeleted-event-visio%28Office.15%29.aspx)|
|[ShapeParentChanged](http://msdn.microsoft.com/library/pages-shapeparentchanged-event-visio%28Office.15%29.aspx)|
|[TextChanged](http://msdn.microsoft.com/library/pages-textchanged-event-visio%28Office.15%29.aspx)|
|[UngroupCanceled](http://msdn.microsoft.com/library/pages-ungroupcanceled-event-visio%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/pages-add-method-visio%28Office.15%29.aspx)|
|[GetNames](http://msdn.microsoft.com/library/pages-getnames-method-visio%28Office.15%29.aspx)|
|[GetNamesU](http://msdn.microsoft.com/library/pages-getnamesu-method-visio%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/pages-application-property-visio%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/pages-count-property-visio%28Office.15%29.aspx)|
|[Document](http://msdn.microsoft.com/library/pages-document-property-visio%28Office.15%29.aspx)|
|[EventList](http://msdn.microsoft.com/library/pages-eventlist-property-visio%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/pages-item-property-visio%28Office.15%29.aspx)|
|[ItemFromID](http://msdn.microsoft.com/library/pages-itemfromid-property-visio%28Office.15%29.aspx)|
|[ItemU](http://msdn.microsoft.com/library/pages-itemu-property-visio%28Office.15%29.aspx)|
|[ObjectType](http://msdn.microsoft.com/library/pages-objecttype-property-visio%28Office.15%29.aspx)|
|[PersistsEvents](http://msdn.microsoft.com/library/pages-persistsevents-property-visio%28Office.15%29.aspx)|
|[Stat](http://msdn.microsoft.com/library/pages-stat-property-visio%28Office.15%29.aspx)|

