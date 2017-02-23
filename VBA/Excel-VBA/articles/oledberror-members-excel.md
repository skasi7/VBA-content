---
title: OLEDBError Members (Excel)
ms.prod: EXCEL
ms.assetid: 52181252-dd6f-b267-fa21-4ad8175b7346
---


# OLEDBError Members (Excel)
Represents an OLE DB error returned by the most recent OLE DB query.

Represents an OLE DB error returned by the most recent OLE DB query.


## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](oledberror-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[Creator](oledberror-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[ErrorString](oledberror-errorstring-property-excel.md)|Returns a  **String** value that represents the ODBC error string.|
|[Native](oledberror-native-property-excel.md)|Returns a provider-specific numeric value that specifies an error. The error number corresponds to an error condition that resulted after the most recent OLE DB query. Read-only  **Long** .|
|[Number](oledberror-number-property-excel.md)|Returns a numeric value that specifies an error. The error number corresponds to a unique trap number corresponding to an error condition that resulted after the most recent OLE DB query. Read-only  **Long** .|
|[Parent](oledberror-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[SqlState](oledberror-sqlstate-property-excel.md)|Returns the SQL state error. Read-only  **String** .|
|[Stage](oledberror-stage-property-excel.md)|Returns a numeric value specifying the stage of an error that resulted after the most recent OLE DB query. Read-only  **Long** .|

