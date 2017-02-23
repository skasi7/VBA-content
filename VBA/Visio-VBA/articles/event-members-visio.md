---
title: Event Members (Visio)
ms.prod: VISIO
ms.assetid: 130f06df-f649-9799-69fc-63f76b530907
---


# Event Members (Visio)
A member of the  **EventList** collection of a source object such as a **Document** . An event encapsulates an event code.

A member of the  **EventList** collection of a source object such as a **Document** . An event encapsulates an event code.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Delete](event-delete-method-visio.md)|Deletes an object.|
|[GetFilterActions](event-getfilteractions-method-visio.md)|Returns an array of the filter actions set for the  **Event** object.|
|[GetFilterCommands](event-getfiltercommands-method-visio.md)|Returns an array of command ranges and a  **True** or **False** value indicating how to filter events for that command range.|
|[GetFilterObjects](event-getfilterobjects-method-visio.md)|Returns an array of object types and a  **True** or **False** value indicating how to filter events for that object.|
|[GetFilterSRC](event-getfiltersrc-method-visio.md)|Returns an array of cell ranges and a  **True** or **False** value indicating whether you are filtering events for that range.|
|[SetFilterActions](event-setfilteractions-method-visio.md)|Specifies the extensions to the  **MouseMove** event that Visio reports.|
|[SetFilterCommands](event-setfiltercommands-method-visio.md)|Specifies an array of command ranges and a  **True** or **False** value indicating how to filter events for each command range.|
|[SetFilterObjects](event-setfilterobjects-method-visio.md)|Specifies an array of object types and a  **True** or **False** value indicating how to filter events for each object.|
|[SetFilterSRC](event-setfiltersrc-method-visio.md)|Specifies an array of cell ranges and a  **True** or **False** value indicating how to filter events for each cell range.|
|[Trigger](event-trigger-method-visio.md)|Causes an event's action to be performed.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Action](event-action-property-visio.md)|Gets or sets the action code of an  **Event** object. Read/write.|
|[Application](event-application-property-visio.md)|Returns the instance of Microsoft Visio that is associated with an object. Read-only.|
|[Enabled](event-enabled-property-visio.md)|Determines whether or not an  **Event** object is currently enabled. Read/write.|
|[Event](event-event-property-visio.md)|Gets or sets the event code of an  **Event** object—an event-action pair. When the event occurs, the action is performed. Read/write.|
|[EventList](event-eventlist-property-visio.md)|Returns the  **EventList** collection of an object or the **EventList** collection that contains an **Event** object. Read-only.|
|[ID](event-id-property-visio.md)|Gets the ID of an object. Read-only.|
|[Index](event-index-property-visio.md)|Gets the ordinal position of an  **Event** object in the **EventList** collection. Read-only.|
|[ObjectType](event-objecttype-property-visio.md)|Returns an object's type. Read-only.|
|[Persistable](event-persistable-property-visio.md)|Determines whether an event can potentially persist within its document. Read-only.|
|[Persistent](event-persistent-property-visio.md)|Determines whether an event persists with its document. Read/write.|
|[Target](event-target-property-visio.md)|Gets or sets the target of an event. Read/write.|
|[TargetArgs](event-targetargs-property-visio.md)|Gets or sets the arguments to be sent to the target of an event. Read/write.|

