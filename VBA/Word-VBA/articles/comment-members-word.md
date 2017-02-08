---
title: Comment Members (Word)
ms.prod: WORD
ms.assetid: 1f1dbb3e-d0ae-9eb7-108a-697a10533e2b
---


# Comment Members (Word)
Represents a single comment. The  **Comment** object is a member of the **[Comments](comments-object-word.md)** collection. The **Comments** collection includes comments in a selection, range or document.

Represents a single comment. The  **Comment** object is a member of the **[Comments](comments-object-word.md)** collection. The **Comments** collection includes comments in a selection, range or document.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[DeleteRecursively](comment-deleterecursively-method-word.md)|Deletes the specified comment and all replies associated with it.|
|[Edit](comment-edit-method-word.md)|Opens the specified OLE object for editing in the application it was created in.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Ancestor](comment-ancestor-property-word.md)|For comments that are replies to existing comments, returns the parent  **Comment** object; for new (top-level) comments, returns null. Read-only.|
|[Application](comment-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[Contact](comment-contact-property-word.md)|Returns a [CoAuthor](coauthor-object-word.md) object that represents the author of the specified comment. Read-only.|
|[Creator](comment-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[Date](comment-date-property-word.md)|Returns a  **Date** that represents the date and time that a comment was inserted. Read-only.|
|[Done](comment-done-property-word.md)|Returns or sets a  **Boolean** whose value is **true** if the specified comment has been marked closed. Read-write.|
|[Index](comment-index-property-word.md)|Returns a  **Long** that represents the position of an item in a collection. Read-only.|
|[IsInk](comment-isink-property-word.md)|Returns a  **Boolean** that represents whether a comment is a handwritten comment.|
|[Parent](comment-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **Comment** object.|
|[Range](comment-range-property-word.md)|Returns a  **Range** object that represents the contents of a comment.|
|[Reference](comment-reference-property-word.md)|Returns a  **Range** object that represents a reference mark for a comment.|
|[Replies](comment-replies-property-word.md)|Returns a [Comments](comments-object-word.md) collection of **Comment** objects that are children of the specified comment. Read-only.|
|[Scope](comment-scope-property-word.md)|Returns a  **[Range](range-object-word.md)** object that represents the range of text marked by the specified comment.|

