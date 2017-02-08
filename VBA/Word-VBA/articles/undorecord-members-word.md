---
title: UndoRecord Members (Word)
ms.prod: WORD
ms.assetid: 50e7d978-f828-d595-9a03-89bd91b14685
---


# UndoRecord Members (Word)
Provides an entry point into the undo stack.

Provides an entry point into the undo stack.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[EndCustomRecord](undorecord-endcustomrecord-method-word.md)|Completes the creation of a custom undo record.|
|[StartCustomRecord](undorecord-startcustomrecord-method-word.md)|Initiates the creation of a custom undo record.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](undorecord-application-property-word.md)|Returns an [Application](application-object-word.md) object that represents the Word application. Read-only.|
|[Creator](undorecord-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[CustomRecordLevel](undorecord-customrecordlevel-property-word.md)|Returns a  **Long** that specifies the number of custom undo action calls that are currently active. Read-only.|
|[CustomRecordName](undorecord-customrecordname-property-word.md)|Returns a  **String** that specifies the entry that appears on the undo stack when all custom undo actions have completed. Read-only.|
|[IsRecordingCustomRecord](undorecord-isrecordingcustomrecord-property-word.md)|Returns a  **Boolean** that specifies whether a custom undo action is being recorded. Read-only.|
|[Parent](undorecord-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified[UndoRecord](undorecord-object-word.md) object.|

