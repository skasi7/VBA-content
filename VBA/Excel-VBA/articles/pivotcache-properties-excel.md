---
title: PivotCache Properties (Excel)
ms.prod: EXCEL
ms.assetid: f075672d-f01b-4fbe-934d-b67395015bda
---


# PivotCache Properties (Excel)

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[ADOConnection](pivotcache-adoconnection-property-excel.md)|Returns an  **ADO Connection** object if the PivotTable cache is connected to an OLE DB data source. The **ADOConnection** property exposes the Microsoft Excel connection to the data provider, allowing the user to write code within the context of the same session that Excel is using with ADO (relational source) or ADO MD (OLAP source). Read-only.|
|[Application](pivotcache-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[BackgroundQuery](pivotcache-backgroundquery-property-excel.md)| **True** if queries for the PivotTable report are performed asynchronously (in the background). Read/write **Boolean** .|
|[CommandText](pivotcache-commandtext-property-excel.md)|Returns or sets the command string for the specified data source. Read/write  **Variant** .|
|[CommandType](pivotcache-commandtype-property-excel.md)|Returns or sets one of the  **[XlCmdType](xlcmdtype-enumeration-excel.md)** constants listed in the following table in the remarks section. The constant that is returned or set describes the value of the **[CommandText](pivotcache-commandtext-property-excel.md)** property. The default value is **xlCmdSQL** . Read/write **XlCmdType** .|
|[Connection](pivotcache-connection-property-excel.md)|Returns or sets a string that contains one of the following: OLE DB settings that enable Microsoft Excel to connect to an OLE DB data source; ODBC settings that enable Microsoft Excel to connect to an ODBC data source; a URL that enables Microsoft Excel to connect to a Web data source; the path to and file name of a text file, or the path to and file name of a file that specifies a database or Web query. Read/write  **Variant** .|
|[Creator](pivotcache-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[EnableRefresh](pivotcache-enablerefresh-property-excel.md)| **True** if the PivotTable cache or query table can be refreshed by the user. The default value is **True** . Read/write **Boolean** .|
|[Index](pivotcache-index-property-excel.md)|Returns a  **Long** value that represents the index number of the object within the collection of similar objects.|
|[IsConnected](pivotcache-isconnected-property-excel.md)|Returns  **True** if the **MaintainConnection** property is **True** and the PivotTable cache is currently connected to its source. Returns **False** if it is not currently connected to its source. Read-only **Boolean** .|
|[LocalConnection](pivotcache-localconnection-property-excel.md)|Returns or sets the connection string to an offline cube file. Read/write  **String** .|
|[MaintainConnection](pivotcache-maintainconnection-property-excel.md)| **True** if the connection to the specified data source is maintained after the refresh and until the workbook is closed. The default value is **True** . Read/write **Boolean** .|
|[MemoryUsed](pivotcache-memoryused-property-excel.md)|Returns the amount of memory currently being used by the object, in bytes. Read-only  **Long** .|
|[MissingItemsLimit](pivotcache-missingitemslimit-property-excel.md)|Returns or sets the maximum quantity of unique items per PivotTable field that are retained even when they have no supporting data in the cache records. Read/write  **[XlPivotTableMissingItems](xlpivottablemissingitems-enumeration-excel.md)** .|
|[OLAP](pivotcache-olap-property-excel.md)|Returns  **True** if the PivotTable cache is connected to an Online Analytical Processing (OLAP) server. Read-only **Boolean** .|
|[OptimizeCache](pivotcache-optimizecache-property-excel.md)| **True** if the PivotTable cache is optimized when it's constructed. The default value is **False** . Read/write **Boolean** .|
|[Parent](pivotcache-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[QueryType](pivotcache-querytype-property-excel.md)|Indicates the type of query used by Microsoft Excel to populate the PivotTable cache. Read-only  **[XlQueryType](xlquerytype-enumeration-excel.md)** .|
|[RecordCount](pivotcache-recordcount-property-excel.md)|Returns the number of records in the PivotTable cache or the number of cache records that contain the specified item. Read-only  **Long** .|
|[Recordset](pivotcache-recordset-property-excel.md)|Returns or sets a  **Recordset** object that's used as the data source for the specified PivotTable cache. Read/write.|
|[RefreshDate](pivotcache-refreshdate-property-excel.md)|Returns the date on which the cache was last refreshed. Read-only  **Date** .|
|[RefreshName](pivotcache-refreshname-property-excel.md)|Returns the name of the person who last refreshed the PivotTable cache. Read-only  **String** .|
|[RefreshOnFileOpen](pivotcache-refreshonfileopen-property-excel.md)| **True** if the PivotTable cache is automatically updated each time the workbook is opened. The default value is **False** . Read/write **Boolean** .|
|[RefreshPeriod](pivotcache-refreshperiod-property-excel.md)|Returns or sets the number of minutes between refreshes. Read/write  **Long** .|
|[RobustConnect](pivotcache-robustconnect-property-excel.md)|Returns or sets how the PivotTable cache connects to its data source. Read/write  **[XlRobustConnect](xlrobustconnect-enumeration-excel.md)** .|
|[SavePassword](pivotcache-savepassword-property-excel.md)| **True** if password information in an ODBC connection string is saved with the specified query. **False** if the password is removed. Read/write **Boolean** .|
|[SourceConnectionFile](pivotcache-sourceconnectionfile-property-excel.md)|Returns or sets a  **String** indicating the Microsoft Office Data Connection file or similar file that was used to create the PivotTable. Read/write.|
|[SourceData](pivotcache-sourcedata-property-excel.md)|Returns the data source for the PivotTable report, as shown in the following table. Read-write  **Variant** .|
|[SourceDataFile](pivotcache-sourcedatafile-property-excel.md)|Returns a  **String** value that indicates the source data file for the cache of the PivotTable.|
|[SourceType](pivotcache-sourcetype-property-excel.md)|Returns a  **[XlPivotTableSourceType](xlpivottablesourcetype-enumeration-excel.md)** value that represents the type of item being published.|
|[UpgradeOnRefresh](pivotcache-upgradeonrefresh-property-excel.md)|Contains information on whether to upgrade the PivotCache and all connected PivotTables on the next refresh. Read/write  **Boolean** .|
|[UseLocalConnection](pivotcache-uselocalconnection-property-excel.md)|Returns  **True** if the **[LocalConnection](pivotcache-localconnection-property-excel.md)** property is used to specify the string that enables Microsoft Excel to connect to a data source. Returns **False** if the connection string specified by the **Connection** property is used. Read/write **Boolean** .|
|[Version](pivotcache-version-property-excel.md)|Returns the version of Microsoft Excel in which the PivotCache was created. Read-only.|
|[WorkbookConnection](pivotcache-workbookconnection-property-excel.md)|Establishes a connection between the current workbook and the  **PivotCache** object. Read-only.|

