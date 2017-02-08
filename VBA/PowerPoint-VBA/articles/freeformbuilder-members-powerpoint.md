---
title: FreeformBuilder Members (PowerPoint)
ms.prod: POWERPOINT
ms.assetid: 2673b640-8aec-1af4-55fd-38d0ad4c9381
---


# FreeformBuilder Members (PowerPoint)
Represents the geometry of a freeform while it is being built.

Represents the geometry of a freeform while it is being built.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[AddNodes](freeformbuilder-addnodes-method-powerpoint.md)|Inserts a new segment at the end of the freeform that's being created, and adds the nodes that define the segment. You can use this method as many times as you want to add nodes to the freeform you're creating. When you finish adding nodes, use the  **[ConvertToShape](freeformbuilder-converttoshape-method-powerpoint.md)** method to create the freeform you've just defined. To add nodes to a freeform after it is been created, use the **[Insert](freeformbuilder-converttoshape-method-powerpoint.md)** method of the **[ShapeNodes](shapenodes-object-powerpoint.md)** collection.|
|[ConvertToShape](freeformbuilder-converttoshape-method-powerpoint.md)|Creates a shape that has the geometric characteristics of the specified  **[FreeformBuilder](freeformbuilder-object-powerpoint.md)** object. Returns a **[Shape](shape-object-powerpoint.md)** object that represents the new shape.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](freeformbuilder-application-property-powerpoint.md)|Returns an  **[Application](application-object-powerpoint.md)** object that represents the creator of the specified object.|
|[Creator](freeformbuilder-creator-property-powerpoint.md)|Returns a  **Long** that represents the four-character creator code for the application in which the specified object was created. For example, if the object was created in Microsoft PowerPoint, this property returns the hexadecimal number 50575054. Read-only.|
|[Parent](freeformbuilder-parent-property-powerpoint.md)|Returns the parent object for the specified object.|

