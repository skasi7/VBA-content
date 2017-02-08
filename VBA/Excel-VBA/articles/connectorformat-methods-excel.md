---
title: ConnectorFormat Methods (Excel)
ms.prod: EXCEL
ms.assetid: d14c80e7-7d14-4d61-a3e5-ab2d0dc14c06
---


# ConnectorFormat Methods (Excel)

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[BeginConnect](connectorformat-beginconnect-method-excel.md)|Attaches the beginning of the specified connector to a specified shape. If there's already a connection between the beginning of the connector and another shape, that connection is broken. If the beginning of the connector isn't already positioned at the specified connecting site, this method moves the beginning of the connector to the connecting site and adjusts the size and position of the connector. Use the  **[EndConnect](connectorformat-endconnect-method-excel.md)** method to attach the end of the connector to a shape.|
|[BeginDisconnect](connectorformat-begindisconnect-method-excel.md)|Detaches the beginning of the specified connector from the shape it's attached to. This method doesn't alter the size or position of the connector: the beginning of the connector remains positioned at a connection site but is no longer connected. Use the  **[EndDisconnect](connectorformat-enddisconnect-method-excel.md)** method to detach the end of the connector from a shape.|
|[EndConnect](connectorformat-endconnect-method-excel.md)|Attaches the end of the specified connector to a specified shape. If there's already a connection between the end of the connector and another shape, that connection is broken. If the end of the connector isn't already positioned at the specified connecting site, this method moves the end of the connector to the connecting site and adjusts the size and position of the connector. Use the  **[BeginConnect](connectorformat-beginconnect-method-excel.md)** method to attach the beginning of the connector to a shape.|
|[EndDisconnect](connectorformat-enddisconnect-method-excel.md)|Detaches the end of the specified connector from the shape it's attached to. This method doesn't alter the size or position of the connector: the end of the connector remains positioned at a connection site but is no longer connected. Use the  **[BeginDisconnect](connectorformat-begindisconnect-method-excel.md)** method to detach the beginning of the connector from a shape.|

