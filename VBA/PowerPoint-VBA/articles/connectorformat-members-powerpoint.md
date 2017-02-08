---
title: ConnectorFormat Members (PowerPoint)
ms.prod: POWERPOINT
ms.assetid: 446eda0c-4992-d38f-b054-355de3058011
---


# ConnectorFormat Members (PowerPoint)
Contains properties and methods that apply to connectors. 

Contains properties and methods that apply to connectors. 


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[BeginConnect](connectorformat-beginconnect-method-powerpoint.md)|Attaches the beginning of the specified connector to a specified shape. |
|[BeginDisconnect](connectorformat-begindisconnect-method-powerpoint.md)|Detaches the beginning of the specified connector from the shape it is attached to. |
|[EndConnect](connectorformat-endconnect-method-powerpoint.md)|Attaches the end of the specified connector to a specified shape. |
|[EndDisconnect](connectorformat-enddisconnect-method-powerpoint.md)|Detaches the end of the specified connector from the shape it is attached to. This method doesn't alter the size or position of the connector: the end of the connector remains positioned at a connection site but is no longer connected. Use the  **[BeginDisconnect](connectorformat-begindisconnect-method-powerpoint.md)** method to detach the beginning of the connector from a shape.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](connectorformat-application-property-powerpoint.md)|Returns an  **[Application](application-object-powerpoint.md)** object that represents the creator of the specified object.|
|[BeginConnected](connectorformat-beginconnected-property-powerpoint.md)|Determines whether the beginning of the specified connector is connected to a shape. Read/write.|
|[BeginConnectedShape](connectorformat-beginconnectedshape-property-powerpoint.md)|Returns a  **[Shape](shape-object-powerpoint.md)** object that represents the shape that the beginning of the specified connector is attached to. Read-only.|
|[BeginConnectionSite](connectorformat-beginconnectionsite-property-powerpoint.md)|Returns an integer that specifies the connection site that the beginning of a connector is connected to. Read-only. |
|[Creator](connectorformat-creator-property-powerpoint.md)|Returns a  **Long** that represents the four-character creator code for the application in which the specified object was created. For example, if the object was created in Microsoft PowerPoint, this property returns the hexadecimal number 50575054. Read-only.|
|[EndConnected](connectorformat-endconnected-property-powerpoint.md)|Determines whether the end of the specified connector is connected to a shape. Read-only.|
|[EndConnectedShape](connectorformat-endconnectedshape-property-powerpoint.md)|Returns a  **[Shape](shape-object-powerpoint.md)** object that represents the shape that the end of the specified connector is attached to. Read-only.|
|[EndConnectionSite](connectorformat-endconnectionsite-property-powerpoint.md)|Returns an integer that specifies the connection site that the end of a connector is connected to. Read-only. |
|[Parent](connectorformat-parent-property-powerpoint.md)|Returns the parent object for the specified object.|
|[Type](connectorformat-type-property-powerpoint.md)|Represents the type of connector. Read/write.|

