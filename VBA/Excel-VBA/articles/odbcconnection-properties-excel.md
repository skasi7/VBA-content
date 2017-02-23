---
title: ODBCConnection Properties (Excel)
ms.prod: EXCEL
ms.assetid: bee9ea5a-4d95-42a2-9f50-68a367ae26a8
---


# ODBCConnection Properties (Excel)

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[AlwaysUseConnectionFile](odbcconnection-alwaysuseconnectionfile-property-excel.md)| **True** if the connection file is always used to establish connection to the data source. Read/write **Boolean** .|
|[Application](odbcconnection-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object. Read-only.|
|[BackgroundQuery](odbcconnection-backgroundquery-property-excel.md)| **True** if queries for the ODBC connection are performed asynchronously (in the background). Read/write **Boolean** .|
|[CommandText](odbcconnection-commandtext-property-excel.md)|Returns or sets the command string for the specified data source. Read/write  **Variant** .|
|[CommandType](odbcconnection-commandtype-property-excel.md)|Returns or sets one of the  **XlCmdType** constants. Read/write **[XlCmdType](xlcmdtype-enumeration-excel.md)** .|
|[Connection](odbcconnection-connection-property-excel.md)|Returns or sets a string that contains ODBC settings that enable Microsoft Excel to connect to an ODBC data source. Read/write  **Variant** .|
|[Creator](odbcconnection-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[EnableRefresh](odbcconnection-enablerefresh-property-excel.md)| **True** if the connection can be refreshed by the user. The default value is **True** . Read/write **Boolean** .|
|[Parent](odbcconnection-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[RefreshDate](odbcconnection-refreshdate-property-excel.md)|Returns the date on which the ODBC connection was last refreshed. Read-only  **Date** .|
|[Refreshing](odbcconnection-refreshing-property-excel.md)| **True** if a background ODBC query is in progress for the specified ODBC connection. Read/write **Boolean** .|
|[RefreshOnFileOpen](odbcconnection-refreshonfileopen-property-excel.md)| **True** if the connection is automatically updated each time the workbook is opened. The default value is **False** . Read/write **Boolean** .|
|[RefreshPeriod](odbcconnection-refreshperiod-property-excel.md)|Returns or sets the number of minutes between refreshes. Read/write  **Long** .|
|[RobustConnect](odbcconnection-robustconnect-property-excel.md)|Returns or sets how ODBC connection connects to its data source. Read/write  **[XlRobustConnect](xlrobustconnect-enumeration-excel.md)** .|
|[SavePassword](odbcconnection-savepassword-property-excel.md)| **True** if password information in an ODBC connection string is saved in the connection string. **False** if the password is removed. Read/write **Boolean** .|
|[ServerCredentialsMethod](odbcconnection-servercredentialsmethod-property-excel.md)|Returns or sets the type of credentials that should be used for server authentication. Read/write  **[XlCredentialsMethod](xlcredentialsmethod-enumeration-excel.md)** .|
|[ServerSSOApplicationID](odbcconnection-serverssoapplicationid-property-excel.md)|Returns or sets a single sign-on application (SSO) identifier that is used to do a lookup in the SSO database for credentials. Read/write  **String** .|
|[SourceConnectionFile](odbcconnection-sourceconnectionfile-property-excel.md)|Returns or sets a  **String** indicating the Microsoft Office Data Connection file or similar file that was used to create the connection. Read/write.|
|[SourceData](odbcconnection-sourcedata-property-excel.md)|Returns the data source for the ODBC connection, as shown in the table. Read/write  **Variant** .|
|[SourceDataFile](odbcconnection-sourcedatafile-property-excel.md)|Returns or sets a  **String** indicating the source data file for an ODBC connection. Read/write.|

