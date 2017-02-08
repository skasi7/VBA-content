---
title: OLEDBConnection Members (Excel)
ms.prod: EXCEL
ms.assetid: 2f1a2f81-ee3a-1b60-8dc3-87818e1790c1
---


# OLEDBConnection Members (Excel)
Represents the OLE DB connection.

Represents the OLE DB connection.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[CancelRefresh](oledbconnection-cancelrefresh-method-excel.md)|Cancels all refresh operations in progress for the specified OLE DB connection.|
|[MakeConnection](oledbconnection-makeconnection-method-excel.md)|Establishes a connection for the specified OLE DB connection.|
|[Reconnect](oledbconnection-reconnect-method-excel.md)|Drops and then reconnects the specified connection.|
|[Refresh](oledbconnection-refresh-method-excel.md)|Refreshes an OLE DB connection.|
|[SaveAsODC](oledbconnection-saveasodc-method-excel.md)|Saves the OLE DB connection as an Microsoft Office Data Connection file.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[ADOConnection](oledbconnection-adoconnection-property-excel.md)|Returns an ADO connection object if the PivotTable cache is connected to an OLE DB data source. Read-only.|
|[AlwaysUseConnectionFile](oledbconnection-alwaysuseconnectionfile-property-excel.md)| **True** if the connection file is always used to establish connection to the data source. Read/write **Boolean** .|
|[Application](oledbconnection-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object. Read-only.|
|[BackgroundQuery](oledbconnection-backgroundquery-property-excel.md)| **True** if queries for the OLE DB connection are performed asynchronously (in the background). Read/write **Boolean** .|
|[CalculatedMembers](oledbconnection-calculatedmembers-property-excel.md)|Returns the  **[CalculatedMembers](calculatedmembers-object-excel.md)** collection for the specified connection. Read-only|
|[CommandText](oledbconnection-commandtext-property-excel.md)|Returns or sets the command string for the specified data source. Read/write  **Variant** .|
|[CommandType](oledbconnection-commandtype-property-excel.md)|Returns or sets one of the  **XlCmdType** constants. Read/write **[XlCmdType](xlcmdtype-enumeration-excel.md)** .|
|[Connection](oledbconnection-connection-property-excel.md)|Returns or sets a string that contains OLE DB settings that enable Microsoft Excel to connect to an OLE DB data source. Read/write  **Variant** .|
|[Creator](oledbconnection-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[EnableRefresh](oledbconnection-enablerefresh-property-excel.md)| **True** if the connection can be refreshed by the user. The default value is **True** . Read/write **Boolean** .|
|[IsConnected](oledbconnection-isconnected-property-excel.md)|Returns  **True** if the **[MaintainConnection](oledbconnection-maintainconnection-property-excel.md)** property is ** True** . Returns **False** if it is not currently connected to its source. Read-only **Boolean** .|
|[LocalConnection](oledbconnection-localconnection-property-excel.md)|Returns or sets the connection string to an offline cube file. Read/write  **String** .|
|[LocaleID](oledbconnection-localeid-property-excel.md)|Returns or sets the locale identifier for the specified connection. Read/write|
|[MaintainConnection](oledbconnection-maintainconnection-property-excel.md)|Returns  **True** if the connection to the specified data source is maintained after the refresh operation and until the workbook is closed. Read/write **Boolean** .|
|[MaxDrillthroughRecords](oledbconnection-maxdrillthroughrecords-property-excel.md)|Returns or sets the maximum number of records to retrieve. Read/write  **Long** .|
|[OLAP](oledbconnection-olap-property-excel.md)|Returns  **True** if the OLE DB connection is connected to an Online Analytical Processing (OLAP) server. Read-only **Boolean** .|
|[Parent](oledbconnection-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[RefreshDate](oledbconnection-refreshdate-property-excel.md)|Returns the date on which the OLE DB connection was last refreshed. Read-only  **Date** .|
|[Refreshing](oledbconnection-refreshing-property-excel.md)| **True** if a background OLE DB query is in progress for the specified OLE DB connection. Read/write **Boolean** .|
|[RefreshOnFileOpen](oledbconnection-refreshonfileopen-property-excel.md)| **True** if the connection is automatically updated each time the workbook is opened. The default value is **False** . Read/write **Boolean** .|
|[RefreshPeriod](oledbconnection-refreshperiod-property-excel.md)|Returns or sets the number of minutes between refreshes. Read/write  **Long** .|
|[RetrieveInOfficeUILang](oledbconnection-retrieveinofficeuilang-property-excel.md)| **True** if the data and errors are to be retrieved in the Office user interface display language when available. Read/write **Boolean** .|
|[RobustConnect](oledbconnection-robustconnect-property-excel.md)| Returns or sets how OLE DB connection connects to its data source. Read/write **[XlRobustConnect](xlrobustconnect-enumeration-excel.md)** .|
|[SavePassword](oledbconnection-savepassword-property-excel.md)| **True** if password information in an OLE DB connection string is saved in the connection string. **False** if the password is removed. Read/write **Boolean** .|
|[ServerCredentialsMethod](oledbconnection-servercredentialsmethod-property-excel.md)|Returns or sets the type of credentials that should be used for server authentication. Read/write  **[XlCredentialsMethod](xlcredentialsmethod-enumeration-excel.md)** .|
|[ServerFillColor](oledbconnection-serverfillcolor-property-excel.md)| **True** if the fill color format for the OLAP server is retrieved from the server when using the specified connection. Read/write **Boolean** .|
|[ServerFontStyle](oledbconnection-serverfontstyle-property-excel.md)| **True** if the font style format for the OLAP server is retrieved from the server when using the specified connection. Read/write **Boolean** .|
|[ServerNumberFormat](oledbconnection-servernumberformat-property-excel.md)| **True** if the number format for the OLAP server is retrieved from the server when using the specified connection. Read/write **Boolean** .|
|[ServerSSOApplicationID](oledbconnection-serverssoapplicationid-property-excel.md)|Returns or sets a single sign-on application (SSO) identifier that is used to perform a lookup in the SSO database for credentials. Read/write  **String** .|
|[ServerTextColor](oledbconnection-servertextcolor-property-excel.md)| **True** if the text color format for the OLAP server is retrieved from the server when using the specified connection. Read/write **Boolean** .|
|[SourceConnectionFile](oledbconnection-sourceconnectionfile-property-excel.md)|Returns or sets a  **String** indicating the Microsoft Office Data Connection file or similar file that was used to create the connection. Read/write.|
|[SourceDataFile](oledbconnection-sourcedatafile-property-excel.md)|Returns or sets a  **String** indicating the source data file for an OLE DB connection. Read/write.|
|[UseLocalConnection](oledbconnection-uselocalconnection-property-excel.md)| **True** if the **[LocalConnection](oledbconnection-localconnection-property-excel.md)** property is used to specify the string that enables Microsoft Excel to connect to a data source. **False** if the connection string specified by the **[Connection](oledbconnection-connection-property-excel.md)** property is used. Read/write **Boolean** .|

