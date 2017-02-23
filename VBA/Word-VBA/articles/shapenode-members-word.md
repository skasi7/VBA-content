---
title: ShapeNode Members (Word)
ms.prod: WORD
ms.assetid: 55803c23-5f6e-aa8c-6e9f-6d350ec71f5e
---


# ShapeNode Members (Word)
Represents the geometry and the geometry-editing properties of the nodes in a user-defined freeform. Nodes include the vertices between the segments of the freeform and the control points for curved segments. The  **ShapeNode** object is a member of the **ShapeNodes** collection. The **[ShapeNodes](shapenodes-object-word.md)** collection contains all the nodes in a freeform.

Represents the geometry and the geometry-editing properties of the nodes in a user-defined freeform. Nodes include the vertices between the segments of the freeform and the control points for curved segments. The  **ShapeNode** object is a member of the **ShapeNodes** collection. The **[ShapeNodes](shapenodes-object-word.md)** collection contains all the nodes in a freeform.


## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](shapenode-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[Creator](shapenode-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[EditingType](shapenode-editingtype-property-word.md)|If the specified node is a vertex, this property returns a value that indicates how changes made to the node affect the two segments connected to the node. Read-only  **MsoEditingType** . .|
|[Parent](shapenode-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **ShapeNode** object.|
|[Points](shapenode-points-property-word.md)|Returns the position of the specified node as a coordinate pair. Read-only  **Variant** .|
|[SegmentType](shapenode-segmenttype-property-word.md)|Returns a value that indicates whether the segment associated with the specified node is straight or curved. Read-only  **MsoSegmentType** .|

