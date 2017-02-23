---
title: Results Members (Outlook)
ms.prod: OUTLOOK
ms.assetid: 650f59fb-0dbd-3f5f-b289-2dfe9e33c20e
---


# Results Members (Outlook)
Contains data and results returned by the  **[Search](search-object-outlook.md)** object and the **[AdvancedSearch](application-advancedsearch-method-outlook.md)** method.

Contains data and results returned by the  **[Search](search-object-outlook.md)** object and the **[AdvancedSearch](application-advancedsearch-method-outlook.md)** method.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[ItemAdd](results-itemadd-event-outlook.md)|Occurs when one or more items are added to the specified collection.|
|[ItemChange](results-itemchange-event-outlook.md)|Occurs when an item in the specified collection is changed.|
|[ItemRemove](results-itemremove-event-outlook.md)|Occurs when an item is deleted from the specified collection.|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[GetFirst](results-getfirst-method-outlook.md)|Returns the first object in the collection.|
|[GetLast](results-getlast-method-outlook.md)|Returns the last object in the collection. |
|[GetNext](results-getnext-method-outlook.md)|Returns the next object in the collection. |
|[GetPrevious](results-getprevious-method-outlook.md)|Returns the previous object in the collection. |
|[Item](results-item-method-outlook.md)|Returns an Outlook item from a collection.|
|[ResetColumns](results-resetcolumns-method-outlook.md)|Clears the properties that have been cached with the  **[SetColumns](results-setcolumns-method-outlook.md)** method.|
|[SetColumns](results-setcolumns-method-outlook.md)|Caches certain properties for extremely fast access to those particular properties of an item within the collection. |
|[Sort](results-sort-method-outlook.md)|Sorts the collection of items by the specified property. The index for the collection is reset to 1 upon completion of this method.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](results-application-property-outlook.md)|Returns an  **[Application](application-object-outlook.md)** object that represents the parent Outlook application for the object. Read-only.|
|[Class](results-class-property-outlook.md)|Returns an  **[OlObjectClass](olobjectclass-enumeration-outlook.md)** constant indicating the object's class. Read-only.|
|[Count](results-count-property-outlook.md)|Returns a  **Long** indicating the count of objects in the specified collection. Read-only.|
|[DefaultItemType](results-defaultitemtype-property-outlook.md)|Returns an  **[OlItemType](olitemtype-enumeration-outlook.md)** constant indicating the default Outlook item type contained in the folder. Read/write.|
|[Parent](results-parent-property-outlook.md)|Returns the parent  **Object** of the specified object. Read-only.|
|[Session](results-session-property-outlook.md)|Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.|

