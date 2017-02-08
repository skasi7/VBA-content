---
title: PivotCell Members (Excel)
ms.prod: EXCEL
ms.assetid: e486cd5d-3f31-29d4-b811-24fc0aed6803
---


# PivotCell Members (Excel)
Represents a cell in a PivotTable report.

Represents a cell in a PivotTable report.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[AllocateChange](pivotcell-allocatechange-method-excel.md)|Performs a writeback operation on the specified cell in a PivotTable report based on an OLAP data source.|
|[DiscardChange](pivotcell-discardchange-method-excel.md)|Discards changes to the specified cell in a PivotTable report.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](pivotcell-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[CellChanged](pivotcell-cellchanged-property-excel.md)|Returns whether a PivotTable value cell has been edited or recalculated since the PivotTable report was created or the last commit operation was performed. Read-only|
|[ColumnItems](pivotcell-columnitems-property-excel.md)|Returns a  **[PivotItemList](pivotitemlist-object-excel.md)** collection that corresponds to the items on the column axis that represent the selected range.|
|[Creator](pivotcell-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[CustomSubtotalFunction](pivotcell-customsubtotalfunction-property-excel.md)|Returns the custom subtotal function field setting of a  **PivotCell** object. Read-only **[XlConsolidationFunction](xlconsolidationfunction-enumeration-excel.md)** .|
|[DataField](pivotcell-datafield-property-excel.md)|Returns a  **[PivotField](pivotfield-object-excel.md)** object that corresponds to the selected data field.|
|[DataSourceValue](pivotcell-datasourcevalue-property-excel.md)|Returns the value last retrieved from the data source for edited cells in a PivotTable report. Read-only|
|[MDX](pivotcell-mdx-property-excel.md)|Returns a tuple that provides the full MDX coordinates of the specified value cell in PivotTable with an OLAP data source. Read-only|
|[Parent](pivotcell-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[PivotCellType](pivotcell-pivotcelltype-property-excel.md)|Returns one of the  **[XlPivotCellType](xlpivotcelltype-enumeration-excel.md)** constants that identifies the PivotTable entity the cell corresponds to. Read-only.|
|[PivotColumnLine](pivotcell-pivotcolumnline-property-excel.md)|Returns the  **PivotLine** on a column for a specific **PivotCell** object. Read-only **PivotLine** .|
|[PivotField](pivotcell-pivotfield-property-excel.md)|Returns a  **[PivotField](pivotfield-object-excel.md)** object that represents the PivotTable field containing the upper-left corner of the specified range.|
|[PivotItem](pivotcell-pivotitem-property-excel.md)|Returns a  **[PivotItem](pivotitem-object-excel.md)** object that represents the PivotTable item containing the upper-left corner of the specified range.|
|[PivotRowLine](pivotcell-pivotrowline-property-excel.md)|Returns the PivotLine on a row for a specific  **PivotCell** object. Read-only **PivotLine** .|
|[PivotTable](pivotcell-pivottable-property-excel.md)|Returns a  **[PivotTable](pivottable-object-excel.md)** object that represents the PivotTable report associated with the PivotCell.|
|[Range](pivotcell-range-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents the range the specified PivotCell applies to.|
|[RowItems](pivotcell-rowitems-property-excel.md)|Returns a  **[PivotItemList](pivotitemlist-object-excel.md)** collection that corresponds to the items on the category axis that represent the selected cell.|
|[ServerActions](pivotcell-serveractions-property-excel.md)|Represents a collection of  _actions_ consisting of OLAP-defined actions which can be executed. The actions are specific to PivotTables existing at a worksheet-level. Read-only|

