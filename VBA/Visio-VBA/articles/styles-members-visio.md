---
title: Styles Members (Visio)
ms.prod: VISIO
ms.assetid: d26e6b65-df20-bcf9-83f0-c44f8998dae0
---


# Styles Members (Visio)
Includes a  **Style** object for each style defined in a document.

Includes a  **Style** object for each style defined in a document.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[BeforeStyleDelete](styles-beforestyledelete-event-visio.md)|Occurs before a style is deleted.|
|[QueryCancelStyleDelete](styles-querycancelstyledelete-event-visio.md)|Occurs before the application deletes a style in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[StyleAdded](styles-styleadded-event-visio.md)|Occurs after a new style is added to a document.|
|[StyleChanged](styles-stylechanged-event-visio.md)|Occurs after the name of a style is changed or a change to the style propagates to objects to which the style is applied.|
|[StyleDeleteCanceled](styles-styledeletecanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelStyleDelete** event.|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Add](styles-add-method-visio.md)|Adds a new  **Style** object to a **Styles** collection.|
|[GetNames](styles-getnames-method-visio.md)|Returns the names of all items in a collection.|
|[GetNamesU](styles-getnamesu-method-visio.md)|Returns the universal names of all items in a collection.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](styles-application-property-visio.md)|Returns the instance of Microsoft Visio that is associated with an object. Read-only.|
|[Count](styles-count-property-visio.md)|Returns the number of objects in a collection. Read-only.|
|[Document](styles-document-property-visio.md)|Gets the  **Document** object that is associated with an object. Read-only.|
|[EventList](styles-eventlist-property-visio.md)|Returns the  **EventList** collection of an object or the **EventList** collection that contains an **Event** object. Read-only.|
|[Item](styles-item-property-visio.md)|Returns an item from a collection. The  **Item** property is the default property for all collections. Read-only.|
|[ItemFromID](styles-itemfromid-property-visio.md)|Returns an item of a collection using the ID of the item. Read-only.|
|[ItemU](styles-itemu-property-visio.md)|Returns an object from a collection. Read-only.|
|[ObjectType](styles-objecttype-property-visio.md)|Returns an object's type. Read-only.|
|[PersistsEvents](styles-persistsevents-property-visio.md)|Indicates whether an object is capable of containing persistent events in its  **EventList** collection. Read-only.|
|[Stat](styles-stat-property-visio.md)|Returns status information for an object. Read-only.|

