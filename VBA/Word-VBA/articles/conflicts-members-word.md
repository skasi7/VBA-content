---
title: Conflicts Members (Word)
ms.prod: WORD
ms.assetid: 395fd60d-6772-9e2a-83b8-562b3c6c6342
---


# Conflicts Members (Word)
 A collection of[Conflict](conflict-object-word.md) objects that represents the conflicts in a document. The type of a **Conflict** object is specified by the[WdRevisionType](wdrevisiontype-enumeration-word.md) enumeration.

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[AcceptAll](conflicts-acceptall-method-word.md)|Accepts all of the user's changes, removes the conflicts, and merges the changes into the server copy of the document.|
|[Item](conflicts-item-method-word.md)|Returns an individual  **Conflicts** object in a collection.|
|[RejectAll](conflicts-rejectall-method-word.md)|Rejects all of the user's changes and retains the server copy of the document.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](conflicts-application-property-word.md)|Returns an [Application](application-object-word.md) object that represents the Microsoft Word application. Read-only.|
|[Count](conflicts-count-property-word.md)|Returns the number of items in the  **Conflicts** collection. Read-only.|
|[Creator](conflicts-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[Parent](conflicts-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **Conflicts** object.|

