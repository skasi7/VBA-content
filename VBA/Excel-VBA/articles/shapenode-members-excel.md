---
title: ShapeNode Members (Excel)
ms.prod: EXCEL
ms.assetid: 76ac3c43-a43f-ee45-2c38-ea237859d03f
---


# ShapeNode Members (Excel)
Represents the geometry and the geometry-editing properties of the nodes in a user-defined freeform.

Represents the geometry and the geometry-editing properties of the nodes in a user-defined freeform.


## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](shapenode-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[Creator](shapenode-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[EditingType](shapenode-editingtype-property-excel.md)|If the specified node is a vertex, this property returns a value that indicates how changes made to the node affect the two segments connected to the node. Read-only  **[MsoEditingType](msoeditingtype-enumeration-office.md)** .|
|[Parent](shapenode-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[Points](shapenode-points-property-excel.md)|Returns the position of the specified node as a coordinate pair. Each coordinate is expressed in points. Read-only  **Variant** .|
|[SegmentType](shapenode-segmenttype-property-excel.md)|Returns a value that indicates whether the segment associated with the specified node is straight or curved. If the specified node is a control point for a curved segment, this property returns  **msoSegmentCurve** . Read-only **MsoSegmentType** .|

