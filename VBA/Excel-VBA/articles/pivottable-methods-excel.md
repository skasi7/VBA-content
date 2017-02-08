---
title: PivotTable Methods (Excel)
ms.prod: EXCEL
ms.assetid: f26eb5ed-6212-4c60-8e88-0753ffd1e84a
---


# PivotTable Methods (Excel)

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[AddDataField](pivottable-adddatafield-method-excel.md)|Adds a data field to a PivotTable report. Returns a  **[PivotField](pivotfield-object-excel.md)** object that represents the new data field.|
|[AddFields](pivottable-addfields-method-excel.md)|Adds row, column, and page fields to a PivotTable report or PivotChart report.|
|[AllocateChanges](pivottable-allocatechanges-method-excel.md)|Performs a writeback operation for all edited cells in a PivotTable report based on an OLAP data source.|
|[CalculatedFields](pivottable-calculatedfields-method-excel.md)|Returns a  **[CalculatedFields](calculatedfields-object-excel.md)** collection that represents all the calculated fields in the specified PivotTable report. Read-only.|
|[ChangeConnection](pivottable-changeconnection-method-excel.md)|Changes the connection of the specified  **[PivotTable](pivottable-object-excel.md)** .|
|[ChangePivotCache](pivottable-changepivotcache-method-excel.md)|Changes the  **[PivotCache](pivotcache-object-excel.md)** of the specified **[PivotTable](pivottable-object-excel.md)** .|
|[ClearAllFilters](pivottable-clearallfilters-method-excel.md)|The  **ClearAllFilters** method deletes all filters currently applied to the PivotTable. This includes deleting all filters in the **PivotFilters** collection of the **PivotTable** object, removing any manual filtering applied and setting all PivotFields in the Report Filter area to the default item.|
|[ClearTable](pivottable-cleartable-method-excel.md)|The  **ClearTable** method is used for clearing a PivotTable. Clearing PivotTables includes removing all the fields and deleting all filtering and sorting applied to the PivotTables. This method resets the PivotTable to the state it had right after it was created, before any fields were added to it.|
|[CommitChanges](pivottable-commitchanges-method-excel.md)|Performs a commit operation on the data source of a PivotTable report based on an OLAP data source.|
|[ConvertToFormulas](pivottable-converttoformulas-method-excel.md)|The  **ConvertToFormulas** method is new in Microsoft Office Excel 2007 and is used for converting a PivotTable to cube formulas. Read/write **Boolean** .|
|[CreateCubeFile](pivottable-createcubefile-method-excel.md)|Creates a cube file from a PivotTable report connected to an Online Analytical Processing (OLAP) data source.|
|[DiscardChanges](pivottable-discardchanges-method-excel.md)|Discards all changes in the edited cells of a PivotTable report based on an OLAP data source.|
|[DrillDown](pivottable-drilldown-method-excel.md)|Enables you to drill down into the data within an OLAP or PowerPivot based cube hierarchy.|
|[DrillTo](pivottable-drillto-method-excel.md)|Enables you to drill to a location within an OLAP or PowerPivot based cube hierarchy.|
|[DrillUp](pivottable-drillup-method-excel.md)|Enables you to drill up into the data within an OLAP or PowerPivot based cube hierarchy.|
|[GetData](pivottable-getdata-method-excel.md)|Returns the value for the a data filed in a PivotTable.|
|[GetPivotData](pivottable-getpivotdata-method-excel.md)|Returns a  **[Range](range-object-excel.md)** object with information about a data item in a PivotTable report.|
|[ListFormulas](pivottable-listformulas-method-excel.md)|Creates a list of calculated PivotTable items and fields on a separate worksheet.|
|[PivotCache](pivottable-pivotcache-method-excel.md)|Returns a  **[PivotCache](pivotcache-object-excel.md)** object that represents the cache for the specified PivotTable report. Read-only.|
|[PivotFields](pivottable-pivotfields-method-excel.md)|Returns an object that represents either a single PivotTable field (a  **[PivotField](comment-object-excel.md)** object) or a collection of both the visible and hidden fields (a **[PivotFields](pivotfields-object-excel.md)** object) in the PivotTable report. Read-only.|
|[PivotSelect](pivottable-pivotselect-method-excel.md)|Selects part of a PivotTable report.|
|[PivotTableWizard](pivottable-pivottablewizard-method-excel.md)|Creates and returns a  **[PivotTable](pivottable-object-excel.md)** object. This method doesn't display the PivotTable Wizard. This method isn't available for OLE DB data sources. Use the **[Add](pivottables-add-method-excel.md)** method to add a PivotTable cache, and then create a PivotTable report based on the cache.|
|[PivotValueCell](pivottable-pivotvaluecell-method-excel.md)|Retrieve the [PivotValueCell Object (Excel)](pivotvaluecell-object-excel.md) object for a given PivotTable provided certain row and column indices.|
|[RefreshDataSourceValues](pivottable-refreshdatasourcevalues-method-excel.md)|Retrieves the current values from the data source for all edited cells in a PivotTable report that is in writeback mode.|
|[RefreshTable](pivottable-refreshtable-method-excel.md)|Refreshes the PivotTable report from the source data. Returns  **True** if it's successful.|
|[RepeatAllLabels](pivottable-repeatalllabels-method-excel.md)|Specifies whether to repeat item labels for all PivotFields in the specified PivotTable.|
|[RowAxisLayout](pivottable-rowaxislayout-method-excel.md)|This method is used for simultaneously setting layout options for all existing PivotFields.|
|[ShowPages](pivottable-showpages-method-excel.md)|Creates a new PivotTable report for each item in the page field. Each new report is created on a new worksheet.|
|[SubtotalLocation](pivottable-subtotallocation-method-excel.md)|This method changes the subtotal location for all existing PivotFields. Changing the subtotal location has an immediate visual effect only for fields in outline form, but it will be set for fields in tabular form as well. |
|[Update](pivottable-update-method-excel.md)|Updates the PivotTable report.|

