---
title: DataFeedConnection Members (Excel)
ms.prod: EXCEL
ms.assetid: 33157c0b-c8d1-355f-8e72-3c7738ff67af
---


# DataFeedConnection Members (Excel)
Contains the data and functionality needed to connect to data feeds. The same object is used for all Data Feed types.

Contains the data and functionality needed to connect to data feeds. The same object is used for all Data Feed types.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[CancelRefresh](datafeedconnection-cancelrefresh-method-excel.md)|Cancels a refresh operation on a data feed connection.|
|[Refresh](datafeedconnection-refresh-method-excel.md)|Refreshes the data feed connection.|
|[SaveAsODC](datafeedconnection-saveasodc-method-excel.md)|Saves the data feed connection as a Microsoft Office Data Connection file.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[AlwaysUseConnectionFile](datafeedconnection-alwaysuseconnectionfile-property-excel.md)| **True** if the connection file is always used to establish connection to the data source. **Boolean** . Read/Write|
|[Application](datafeedconnection-application-property-excel.md)|Returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. Read-only.|
|[CommandText](datafeedconnection-commandtext-property-excel.md)|Returns or sets the command string for the specified data source.  **Variant** Read/Write|
|[CommandType](datafeedconnection-commandtype-property-excel.md)|Returns or sets the command string for the specified data source.  **Variant** Read/Write|
|[Connection](datafeedconnection-connection-property-excel.md)|Returns or sets a string that contains Service Contract settings that enable Microsoft Excel to connect to a Data Feed data source.  **Variant** Read/Write|
|[Creator](datafeedconnection-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which the specified object was created.  **Long** Read-only|
|[EnableRefresh](datafeedconnection-enablerefresh-property-excel.md)| **True** if the connection can be refreshed by the user. The default value is **True** . **Boolean** Read/Write|
|[Parent](datafeedconnection-parent-property-excel.md)|Returns an  **Object** that represents the parent object of the specified[DataFeedConnection Object (Excel)](datafeedconnection-object-excel.md) object. Read-only.|
|[RefreshDate](datafeedconnection-refreshdate-property-excel.md)|Returns the date on which the OLE DB connection was last refreshed.  **Date** . Read-only|
|[Refreshing](datafeedconnection-refreshing-property-excel.md)| **True** if an OLE DB query is in progress for the specified data source connection. **Boolean** Read/Write|
|[RefreshOnFileOpen](datafeedconnection-refreshonfileopen-property-excel.md)| **True** if the connection is automatically updated each time the workbook is opened. The default value is **False** .|
|[RefreshPeriod](datafeedconnection-refreshperiod-property-excel.md)|Returns or sets the number of minutes between refreshes.  **Long** Read/Write|
|[SavePassword](datafeedconnection-savepassword-property-excel.md)| **True** if password information in a data feed connection string is saved in the connection string. **False** if the password is removed.|
|[ServerCredentialsMethod](datafeedconnection-servercredentialsmethod-property-excel.md)|Returns or sets the type of credentials that should be used for server authentication.  **[XlCredentialsMethod Enumeration (Excel)](xlcredentialsmethod-enumeration-excel.md)** Read/Write|
|[SourceConnectionFile](datafeedconnection-sourceconnectionfile-property-excel.md)|Returns or sets a  **String** indicating the Microsoft Office Data Connection file or similar file that was used to create the connection. Read/Write|
|[SourceDataFile](datafeedconnection-sourcedatafile-property-excel.md)|A path to the original file used to create the connection. In the case of an OData connection, this is the location of the *.atom or *.atomsvc file used to create the connection.  **String** Read/Write|

