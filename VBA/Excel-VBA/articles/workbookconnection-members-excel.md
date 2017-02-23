---
title: WorkbookConnection Members (Excel)
ms.prod: EXCEL
ms.assetid: 1c692856-1ddb-1d7d-4463-143cba3dfbe8
---


# WorkbookConnection Members (Excel)
A connection is a set of information needed to obtain data from an external data source other than an Microsoft Office Excel 2007 workbook. 

A connection is a set of information needed to obtain data from an external data source other than an Microsoft Office Excel 2007 workbook. 


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Delete](workbookconnection-delete-method-excel.md)|Deletes a workbook connection.|
|[Refresh](workbookconnection-refresh-method-excel.md)|Refreshes a workbook connection.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](workbookconnection-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object. Read-only.|
|[Creator](workbookconnection-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[DataFeedConnection](workbookconnection-datafeedconnection-property-excel.md)|Returns a  **DataFeedConnection** object that contains the data and functionality needed to connect to data feeds. Read-only|
|[Description](workbookconnection-description-property-excel.md)|Returns or sets a brief description for a  **WorkbookConnection** object. Read/write **String** .|
|[InModel](workbookconnection-inmodel-property-excel.md)|Specifies whether or not the [WorkbookConnection Object (Excel)](workbookconnection-object-excel.md) has been added to the model. **Boolean** Read-only|
|[ModelConnection](workbookconnection-modelconnection-property-excel.md)|Returns an object that contains information for the new Model Connection Type introduced in Excel 2013 to interact with the integrated Data Model. Read-only|
|[ModelTables](workbookconnection-modeltables-property-excel.md)|Returns a [ModelTables Object (Excel)](modeltables-object-excel.md) object associated with the particular connection. Read-only|
|[Name](workbookconnection-name-property-excel.md)|Returns or sets the name of the workbook connection object. Read/write  **String** .|
|[ODBCConnection](workbookconnection-odbcconnection-property-excel.md)|Retuns the ODBC Connection details for the specified  **WorkbookConnection** object. Read-only **ODBCConnection** .|
|[OLEDBConnection](workbookconnection-oledbconnection-property-excel.md)|Retuns the OLEDB Connection details for the specified  **WorkbookConnection** object. Read-only ** OLEDBConnection** .|
|[Parent](workbookconnection-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[Ranges](workbookconnection-ranges-property-excel.md)|Returns the range of object for the specified  **WorkbookConnection** object. Read-only **Ranges** .|
|[RefreshWithRefreshAll](workbookconnection-refreshwithrefreshall-property-excel.md)|Determines if the connection should be refreshed when refresh all is executed.  **Boolean** . Read/Write|
|[TextConnection](workbookconnection-textconnection-property-excel.md)|Returns a [TextConnection Object (Excel)](textconnection-object-excel.md) object that contains the information on a query to a text file. Read-only|
|[Type](workbookconnection-type-property-excel.md)|Returns the workbook connection type. Read-only  **[XlConnectionType](xlconnectiontype-enumeration-excel.md)** .|
|[WorksheetDataConnection](workbookconnection-worksheetdataconnection-property-excel.md)|Returns an object that contains information for a connection from the PowerPivot Model to data within the workbook such as a range, named range, or table. Read-only|

