---
title: ShapeNodes Members (Excel)
ms.prod: EXCEL
ms.assetid: 3964c044-89e0-fb12-16c3-759a63248a24
---


# ShapeNodes Members (Excel)
A collection of all the  **[ShapeNode](shapenode-object-excel.md)** objects in the specified freeform.

A collection of all the  **[ShapeNode](shapenode-object-excel.md)** objects in the specified freeform.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Delete](shapenodes-delete-method-excel.md)|Deletes the object.|
|[Insert](shapenodes-insert-method-excel.md)|Inserts a node into a freeform shape.|
|[Item](shapenodes-item-method-excel.md)|Returns a single object from a collection.|
|[SetEditingType](shapenodes-seteditingtype-method-excel.md)|Sets the editing type of the node specified by  _Index_. If the node is a control point for a curved segment, this method sets the editing type of the node adjacent to it that joins two segments. Note that, depending on the editing type, this method may affect the position of adjacent nodes.|
|[SetPosition](shapenodes-setposition-method-excel.md)|Sets the location of the node specified by  _Index_. Note that, depending on the editing type of the node, this method may affect the position of adjacent nodes.|
|[SetSegmentType](shapenodes-setsegmenttype-method-excel.md)|Sets the segment type of the segment that follows the node specified by  _Index_. If the node is a control point for a curved segment, this method sets the segment type for that curve. Note that this may affect the total number of nodes by inserting or deleting adjacent nodes.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](shapenodes-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[Count](shapenodes-count-property-excel.md)|Returns an  **Integer** value that represents the number of objects in the collection.|
|[Creator](shapenodes-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[Parent](shapenodes-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|

