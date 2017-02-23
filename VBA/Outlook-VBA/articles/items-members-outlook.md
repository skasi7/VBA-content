---
title: Items Members (Outlook)
ms.prod: OUTLOOK
ms.assetid: bcc2cf6c-b6fb-e1a2-1d5c-d7e2bdf6b7dc
---


# Items Members (Outlook)
Contains a collection of [Outlook item objects](outlook-item-objects.md) in a folder.

Contains a collection of [Outlook item objects](outlook-item-objects.md) in a folder.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[ItemAdd](items-itemadd-event-outlook.md)|Occurs when one or more items are added to the specified collection. This event does not run when a large number of items are added to the folder at once. This event is not available in Microsoft Visual Basic Scripting Edition (VBScript).|
|[ItemChange](items-itemchange-event-outlook.md)|Occurs when an item in the specified collection is changed. This event is not available in Microsoft Visual Basic Scripting Edition (VBScript).|
|[ItemRemove](items-itemremove-event-outlook.md)|Occurs when an item is deleted from the specified collection.|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Add](items-add-method-outlook.md)|Creates a new Outlook item in the  **[Items](items-object-outlook.md)** collection for the folder.|
|[Find](items-find-method-outlook.md)|Locates and returns a Microsoft Outlook item object that satisfies the given  _Filter_ .|
|[FindNext](items-findnext-method-outlook.md)|After the  **[Find](items-find-method-outlook.md)** method runs, this method finds and returns the next Outlook item in the specified collection.|
|[GetFirst](items-getfirst-method-outlook.md)|Returns the first object in the collection. |
|[GetLast](items-getlast-method-outlook.md)|Returns the last object in the collection. |
|[GetNext](items-getnext-method-outlook.md)|Returns the next object in the collection. |
|[GetPrevious](items-getprevious-method-outlook.md)|Returns the previous object in the collection. |
|[Item](items-item-method-outlook.md)|Returns an Outlook item from a collection.|
|[Remove](items-remove-method-outlook.md)|Removes an object from the collection.|
|[ResetColumns](items-resetcolumns-method-outlook.md)|Clears the properties that have been cached with the  **[SetColumns](items-setcolumns-method-outlook.md)** method.|
|[Restrict](items-restrict-method-outlook.md)|Applies a filter to the  **[Items](items-object-outlook.md)** collection, returning a new collection containing all of the items from the original that match the filter.|
|[SetColumns](items-setcolumns-method-outlook.md)|Caches certain properties for extremely fast access to those particular properties of each item in an  **[Items](items-object-outlook.md)** collection.|
|[Sort](items-sort-method-outlook.md)|Sorts the collection of items by the specified property. The index for the collection is reset to 1 upon completion of this method.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](items-application-property-outlook.md)|Returns an  **[Application](application-object-outlook.md)** object that represents the parent Outlook application for the object. Read-only.|
|[Class](items-class-property-outlook.md)|Returns an  **[OlObjectClass](olobjectclass-enumeration-outlook.md)** constant indicating the object's class. Read-only.|
|[Count](items-count-property-outlook.md)|Returns a  **Long** indicating the count of objects in the specified collection. Read-only.|
|[IncludeRecurrences](items-includerecurrences-property-outlook.md)|Returns a  **Boolean** that indicates **True** if the **[Items](items-object-outlook.md)** collection should include recurrence patterns. Read/write.|
|[Parent](items-parent-property-outlook.md)|Returns the parent  **Object** of the specified object. Read-only.|
|[Session](items-session-property-outlook.md)|Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.|

