---
title: Revision Members (Word)
ms.prod: WORD
ms.assetid: 97eb185c-125a-1c5f-6f54-157fd5bbf355
---


# Revision Members (Word)
Represents a change marked with a revision mark. The  **Revision** object is a member of the **[Revisions](revisions-object-word.md)** collection. The **Revisions** collection includes all the revision marks in a range or document.

Represents a change marked with a revision mark. The  **Revision** object is a member of the **[Revisions](revisions-object-word.md)** collection. The **Revisions** collection includes all the revision marks in a range or document.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Accept](revision-accept-method-word.md)|Accepts the specified tracked change, removes the revision mark, and incorporates the change into the document.|
|[Reject](revision-reject-method-word.md)|Rejects the specified tracked change. The revision marks are removed, leaving the original text intact.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](revision-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[Author](revision-author-property-word.md)|Returns the name of the user who made the specified tracked change. Read-only  **String** .|
|[Cells](revision-cells-property-word.md)|Returns a  **Cells** collection that represents the table cells that have been marked with revision marks. Read-only.|
|[Creator](revision-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[Date](revision-date-property-word.md)|The date and time that the tracked change was made. Read-only  **Date** .|
|[FormatDescription](revision-formatdescription-property-word.md)|Returns a  **String** representing a description of tracked formatting changes in a revision. Read-only.|
|[Index](revision-index-property-word.md)|Returns a  **Long** that represents the position of an item in a collection. Read-only.|
|[MovedRange](revision-movedrange-property-word.md)|Returns a  **[Range](range-object-word.md)** object that represents a range of text that was moved from one place to another in a document with tracked changes. Read-only.|
|[Parent](revision-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **Revision** object.|
|[Range](revision-range-property-word.md)|Returns a  **Range** object that represents the portion of a document that's contained within a revision mark.|
|[Style](revision-style-property-word.md)|Returns a  **Style** object that represents the style associated with a revision mark.|
|[Type](revision-type-property-word.md)|Returns the revision type. Read-only  **[WdRevisionType](wdrevisiontype-enumeration-word.md)** .|

