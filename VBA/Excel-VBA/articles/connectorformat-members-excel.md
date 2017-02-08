---
title: ConnectorFormat Members (Excel)
ms.prod: EXCEL
ms.assetid: b7597f8e-5f21-c1d6-2b31-9067dd0cc029
---


# ConnectorFormat Members (Excel)
Contains properties and methods that apply to connectors.

Contains properties and methods that apply to connectors.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[BeginConnect](connectorformat-beginconnect-method-excel.md)|Attaches the beginning of the specified connector to a specified shape. If there's already a connection between the beginning of the connector and another shape, that connection is broken. If the beginning of the connector isn't already positioned at the specified connecting site, this method moves the beginning of the connector to the connecting site and adjusts the size and position of the connector. Use the  **[EndConnect](connectorformat-endconnect-method-excel.md)** method to attach the end of the connector to a shape.|
|[BeginDisconnect](connectorformat-begindisconnect-method-excel.md)|Detaches the beginning of the specified connector from the shape it's attached to. This method doesn't alter the size or position of the connector: the beginning of the connector remains positioned at a connection site but is no longer connected. Use the  **[EndDisconnect](connectorformat-enddisconnect-method-excel.md)** method to detach the end of the connector from a shape.|
|[EndConnect](connectorformat-endconnect-method-excel.md)|Attaches the end of the specified connector to a specified shape. If there's already a connection between the end of the connector and another shape, that connection is broken. If the end of the connector isn't already positioned at the specified connecting site, this method moves the end of the connector to the connecting site and adjusts the size and position of the connector. Use the  **[BeginConnect](connectorformat-beginconnect-method-excel.md)** method to attach the beginning of the connector to a shape.|
|[EndDisconnect](connectorformat-enddisconnect-method-excel.md)|Detaches the end of the specified connector from the shape it's attached to. This method doesn't alter the size or position of the connector: the end of the connector remains positioned at a connection site but is no longer connected. Use the  **[BeginDisconnect](connectorformat-begindisconnect-method-excel.md)** method to detach the beginning of the connector from a shape.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](connectorformat-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[BeginConnected](connectorformat-beginconnected-property-excel.md)| **True** if the beginning of the specified connector is connected to a shape. Read-only **[MsoTriState](msotristate-enumeration-office.md)** .|
|[BeginConnectedShape](connectorformat-beginconnectedshape-property-excel.md)|Returns a  **[Shape](shape-object-excel.md)** object that represents the shape that the beginning of the specified connector is attached to. Read-only.|
|[BeginConnectionSite](connectorformat-beginconnectionsite-property-excel.md)|Returns an integer that specifies the connection site that the beginning of a connector is connected to. Read-only  **Long** .|
|[Creator](connectorformat-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[EndConnected](connectorformat-endconnected-property-excel.md)| **msoTrue** if the end of the specified connector is connected to a shape. Read-only **[MsoTriState](msotristate-enumeration-office.md)** .|
|[EndConnectedShape](connectorformat-endconnectedshape-property-excel.md)|Returns a  **[Shape](shape-object-excel.md)** object that represents the shape that the end of the specified connector is attached to. Read-only.|
|[EndConnectionSite](connectorformat-endconnectionsite-property-excel.md)|Returns an integer that specifies the connection site that the end of a connector is connected to. Read-only  **Long** .|
|[Parent](connectorformat-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[Type](connectorformat-type-property-excel.md)|Returns or sets a  **[MsoConnectorType](msoconnectortype-enumeration-office.md)** value that represents the connector format type.|

