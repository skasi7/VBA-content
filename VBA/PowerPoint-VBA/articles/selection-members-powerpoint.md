---
title: Selection Members (PowerPoint)
ms.prod: POWERPOINT
ms.assetid: cfc57277-8872-4d39-0cc7-3d52d514406c
---


# Selection Members (PowerPoint)
Represents the selection in the specified document window. The  **Selection** object is deleted whenever you change slides in an active slide view (the **Type** property will return **ppSelectionNone** ).

Represents the selection in the specified document window. The  **Selection** object is deleted whenever you change slides in an active slide view (the **Type** property will return **ppSelectionNone** ).


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Copy](selection-copy-method-powerpoint.md)|Copies the specified object to the Clipboard.|
|[Cut](selection-cut-method-powerpoint.md)|Deletes the specified object and places it on the Clipboard.|
|[Delete](selection-delete-method-powerpoint.md)|Deletes the specified  **Selection** object.|
|[Unselect](selection-unselect-method-powerpoint.md)|Cancels the current selection.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](selection-application-property-powerpoint.md)|Returns an  **[Application](application-object-powerpoint.md)** object that represents the creator of the specified object.|
|[ChildShapeRange](selection-childshaperange-property-powerpoint.md)|Returns a  **[ShapeRange](shaperange-object-powerpoint.md)** object that represents the child shapes of a selection.|
|[HasChildShapeRange](selection-haschildshaperange-property-powerpoint.md)|**True** if the selection contains child shapes. Read-only.|
|[Parent](selection-parent-property-powerpoint.md)|Returns the parent object for the specified object.|
|[ShapeRange](selection-shaperange-property-powerpoint.md)|Returns a  **[ShapeRange](shaperange-object-powerpoint.md)** object that represents all the slide objects that have been selected on the specified slide. Read-only.|
|[SlideRange](selection-sliderange-property-powerpoint.md)|Returns a  **[SlideRange](sliderange-object-powerpoint.md)** object that represents a range of selected slides. Read-only.|
|[TextRange](selection-textrange-property-powerpoint.md)|Returns a  **[TextRange](textrange-object-powerpoint.md)** object that represents the selected text. Read-only.|
|[TextRange2](selection-textrange2-property-powerpoint.md)|Returns the  **TextRange2** object of the current **Selection** object. Read-only.|
|[Type](selection-type-property-powerpoint.md)|Represents the type of objects in a selection. Read-only.|

