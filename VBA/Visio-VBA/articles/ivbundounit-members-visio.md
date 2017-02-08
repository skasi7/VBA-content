---
title: IVBUndoUnit Members (Visio)
ms.prod: VISIO
ms.assetid: cc17cb26-6c27-d5f8-f535-c93c57b375d4
---


# IVBUndoUnit Members (Visio)
The interface on an undo unit in Microsoft Visio. An undo unit encapsulates the information necessary to undo or redo a single action.

The interface on an undo unit in Microsoft Visio. An undo unit encapsulates the information necessary to undo or redo a single action.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Do](ivbundounit-do-method-visio.md)|Called by the Undo Manager to tell an undo unit to perform its action.|
|[OnNextAdd](ivbundounit-onnextadd-method-visio.md)|Notifies an undo unit that another undo unit has been added to the undo stack. Returns  **Nothing** .|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Description](ivbundounit-description-property-visio.md)|Gets the description of an  **IVBUndoUnit** interface. Read-only.|
|[UnitSize](ivbundounit-unitsize-property-visio.md)|Returns the size of the undo unit in memory, in bytes. Read-only.|
|[UnitTypeCLSID](ivbundounit-unittypeclsid-property-visio.md)|Identifies an undo unit by its class ID (CLSID). Read-only.|
|[UnitTypeLong](ivbundounit-unittypelong-property-visio.md)|Identifies an undo unit by a  **Long** . Read-only.|

