---
title: TableObject Properties (Excel)
ms.prod: EXCEL
ms.assetid: 4ea0d734-bad9-37d0-6e10-cbd7b894223a
---


# TableObject Properties (Excel)

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[AdjustColumnWidth](tableobject-adjustcolumnwidth-property-excel.md)|Specifies if the column widths are automatically adjusted for the best fit each time you refresh the specified query table. The default value is  **True** . **Boolean** Read/Write|
|[Application](tableobject-application-property-excel.md)|Returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. Read-only.|
|[Creator](tableobject-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[Destination](tableobject-destination-property-excel.md)|Returns the cell in the upper-left corner of the query table destination range (the range where the resulting query table will be placed). The destination range must be on the worksheet that contains the  **TableObject** object. **Range** . Read-only|
|[EnableEditing](tableobject-enableediting-property-excel.md)| **True** if the user can edit the specified query table. **False** if the user can only refresh the query table. **Boolean** Read/Write|
|[EnableRefresh](tableobject-enablerefresh-property-excel.md)|Specifies if the query table can be refreshed by the user.  **Boolean** Read/Write|
|[FetchedRowOverflow](tableobject-fetchedrowoverflow-property-excel.md)|Specifies if the number of rows returned by the last use of the Refresh method is greater than the number of rows available on the worksheet.  **Boolean** Read-only|
|[ListObject](tableobject-listobject-property-excel.md)|Returns a [ListObject Object (Excel)](listobject-object-excel.md) object for the[TableObject Object (Excel)](tableobject-object-excel.md) object. Read-only|
|[Parent](tableobject-parent-property-excel.md)|Returns an  **Object** that represents the parent object of the specified[TableObject Object (Excel)](tableobject-object-excel.md) object. Read-only.|
|[PreserveColumnInfo](tableobject-preservecolumninfo-property-excel.md)|Specifies if column sorting, filtering, and layout information is preserved whenever a query table is refreshed. The default value is  **False** . **Boolean** Read/Write|
|[PreserveFormatting](tableobject-preserveformatting-property-excel.md)| **True** if any formatting common to the first five rows of data are applied to new rows of data in the query table. Unused cells aren't formatted. The property is **False** if the last AutoFormat applied to the query table is applied to new rows of data. The default value is **True** . **Boolean** Read/Write|
|[RefreshStyle](tableobject-refreshstyle-property-excel.md)|Returns or sets the way rows on the specified worksheet are added or deleted to accommodate the number of rows in a record set returned by a query. [XlCellInsertionMode Enumeration (Excel)](xlcellinsertionmode-enumeration-excel.md) Read/Write|
|[ResultRange](tableobject-resultrange-property-excel.md)|Returns a [Range Object (Excel)](range-object-excel.md) object that represents the area of the worksheet occupied by the specified query table. Read-only|
|[RowNumbers](tableobject-rownumbers-property-excel.md)|Specifies if row numbers are added as the first column of the specified query table.  **Boolean** Read/Write|
|[WorkbookConnection](tableobject-workbookconnection-property-excel.md)|Returns the [WorkbookConnection Object (Excel)](workbookconnection-object-excel.md) used by the **TableObject** for connecting to the model. The[WorkbookConnection.ModelConnection Property (Excel)](workbookconnection-modelconnection-property-excel.md) property on the **WorkbookConnection** object can then be used to get to and edit DAX etc. Read-only|

