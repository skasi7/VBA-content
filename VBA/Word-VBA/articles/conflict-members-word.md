---
title: Conflict Members (Word)
ms.prod: WORD
ms.assetid: f097cddc-b78a-d154-0b88-ed22a876d946
---


# Conflict Members (Word)
Represents a conflicting edit in a co-authored document. The type of a  **Conflict** object is specified by the[WdRevisionType](wdrevisiontype-enumeration-word.md) enumeration.

Represents a conflicting edit in a co-authored document. The type of a  **Conflict** object is specified by the[WdRevisionType](wdrevisiontype-enumeration-word.md) enumeration.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Accept](conflict-accept-method-word.md)|Accepts the user specified conflict change, and removes the conflict.|
|[Reject](conflict-reject-method-word.md)|Rejects the user change, removes the conflict, and accepts the server copy of the change for the conflict.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](conflict-application-property-word.md)|Returns an [Application](application-object-word.md) object that represents the Microsoft Word application. Read-only.|
|[Creator](conflict-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[Index](conflict-index-property-word.md)|Returns a  **Long** that represents the position of an item in a collection. Read-only.|
|[Parent](conflict-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **Conflict** object.|
|[Range](conflict-range-property-word.md)| Returns a[Range](range-object-word.md) object that represents the portion of a document that is contained in the specified object. Read-only.|
|[Type](conflict-type-property-word.md)|Returns the [WdRevisionType](wdrevisiontype-enumeration-word.md)for the [Conflict](conflict-object-word.md) object. Read-only.|

