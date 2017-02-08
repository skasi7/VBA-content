---
title: PivotTable Properties (Excel)
ms.prod: EXCEL
ms.assetid: 4fe73302-9de4-493f-a52a-276572853c75
---


# PivotTable Properties (Excel)

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[ActiveFilters](pivottable-activefilters-property-excel.md)|Indicates the currently active filter in the specified PivotTable. Read-only.|
|[Allocation](pivottable-allocation-property-excel.md)|Returns or sets whether to run an  **UPDATE CUBE** statement for each cell is edited, or only when the user chooses to calculate changes when performing what-if analysis on a PivotTable based on an OLAP data source. Read/write|
|[AllocationMethod](pivottable-allocationmethod-property-excel.md)|Returns or sets what method to use to allocate values when performing what-if analysis on a PivotTable report based on an OLAP data source. Read/write|
|[AllocationValue](pivottable-allocationvalue-property-excel.md)|Returns or sets what value to allocate when performing what-if analysis on a PivotTable report based on an OLAP data source. Read/write|
|[AllocationWeightExpression](pivottable-allocationweightexpression-property-excel.md)|Returns or sets the MDX weight expression to use when performing what-if analysis on a PivotTable report based on an OLAP data source. Read/write|
|[AllowMultipleFilters](pivottable-allowmultiplefilters-property-excel.md)|Sets or retrieves a value that indicates whether a PivotField can have multiple filters applied to it at the same time. Read/write  **Boolean** .|
|[AlternativeText](pivottable-alternativetext-property-excel.md)|Returns or sets the descriptive (alternative) text string for the specified PivotTable. Read/write|
|[Application](pivottable-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[CacheIndex](pivottable-cacheindex-property-excel.md)|Returns or sets the index number of the PivotTable cache. Read/write  **Long** .|
|[CalculatedMembers](pivottable-calculatedmembers-property-excel.md)|Returns a  **[CalculatedMembers](calculatedmembers-object-excel.md)** collection representing all the calculated members and calculated measures for an OLAP PivotTable.|
|[CalculatedMembersInFilters](pivottable-calculatedmembersinfilters-property-excel.md)|Returns or sets whether to evaluate calculated members from OLAP servers in filters. Read/write|
|[ChangeList](pivottable-changelist-property-excel.md)|Returns the  **[PivotTableChangeList](pivottablechangelist-object-excel.md)** collection that represents the list of changes that have been made to the specified PivotTable based on an OLAP data source. Read-only|
|[ColumnFields](pivottable-columnfields-property-excel.md)|Returns an object that represents either a single PivotTable field (a  **[PivotField](pivotfield-object-excel.md)** object) or a collection of all the fields (a **[PivotFields](pivotfields-object-excel.md)** object) that are currently shown as column fields. Read-only.|
|[ColumnGrand](pivottable-columngrand-property-excel.md)| **True** if the PivotTable report shows grand totals for columns. Read/write **Boolean** .|
|[ColumnRange](pivottable-columnrange-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents the range that contains the column area in the PivotTable report. Read-only.|
|[CompactLayoutColumnHeader](pivottable-compactlayoutcolumnheader-property-excel.md)|Specifies the caption that is displayed in the column header of a PivotTable when in compact row layout form. Read/write  **String** .|
|[CompactLayoutRowHeader](pivottable-compactlayoutrowheader-property-excel.md)|Specifies the caption that is displayed in the row header of a PivotTable when in compact row layout form. Read/write  **String** .|
|[CompactRowIndent](pivottable-compactrowindent-property-excel.md)|Returns or sets the indent increment for PivotItems when compact row layout form is turned on. Read/write.|
|[Creator](pivottable-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[CubeFields](pivottable-cubefields-property-excel.md)|Returns the  **[CubeFields](cubefields-object-excel.md)** collection. Each **[CubeField](cubefield-object-excel.md)** object contains the properties of the cube field element. Read-only.|
|[DataBodyRange](pivottable-databodyrange-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents the range of values in a PivotTable. Read-only.|
|[DataFields](pivottable-datafields-property-excel.md)|Returns an object that represents either a single PivotTable field (a  **[PivotField](pivotfield-object-excel.md)** object) or a collection of all the fields (a **[PivotFields](pivotfields-object-excel.md)** object) that are currently shown as data fields. Read-only.|
|[DataLabelRange](pivottable-datalabelrange-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents the range that contains the labels for the data fields in the PivotTable report. Read-only.|
|[DataPivotField](pivottable-datapivotfield-property-excel.md)|Returns a  **[PivotField](pivotfield-object-excel.md)** object that represents all the data fields in a PivotTable. Read-only.|
|[DisplayContextTooltips](pivottable-displaycontexttooltips-property-excel.md)|Controls whether or not tooltips are displayed for PivotTable cells. Read/write  **Boolean** .|
|[DisplayEmptyColumn](pivottable-displayemptycolumn-property-excel.md)|Returns  **True** when the non-empty MDX keyword is included in the query to the OLAP provider for the value axis. The OLAP provider will not return empty columns in the result set. Returns **False** when the non-empty keyword is omitted. Read/write **Boolean** .|
|[DisplayEmptyRow](pivottable-displayemptyrow-property-excel.md)|Returns  **True** when the non-empty MDX keyword is included in the query to the OLAP provider for the category axis. The OLAP provider will not return empty rows in the result set. Returns **False** when the non-empty keyword is omitted. Read/write **Boolean** .|
|[DisplayErrorString](pivottable-displayerrorstring-property-excel.md)| **True** if the PivotTable report displays a custom error string in cells that contain errors. The default value is **False** . Read/write **Boolean** .|
|[DisplayFieldCaptions](pivottable-displayfieldcaptions-property-excel.md)|Controls whether or not filter buttons and PivotField captions for rows and columns are displayed in the grid. Read/write.|
|[DisplayImmediateItems](pivottable-displayimmediateitems-property-excel.md)|Returns or sets a  **Boolean** that indicates whether items in the row and column areas are visible when the data area of the PivotTable is empty. Set this property to **False** to hide the items in the row and column areas when the data area of the PivotTable is empty. The default value is **True** .|
|[DisplayMemberPropertyTooltips](pivottable-displaymemberpropertytooltips-property-excel.md)|Controls whether or not to display member properties in tooltips. Read/write  **Boolean** .|
|[DisplayNullString](pivottable-displaynullstring-property-excel.md)| **True** if the PivotTable report displays a custom string in cells that contain null values. The default value is **True** . Read/write **Boolean** .|
|[EnableDataValueEditing](pivottable-enabledatavalueediting-property-excel.md)| **True** to disable the alert for when the user overwrites values in the data area of the PivotTable. **True** also allows the user to change data values that previously could not be changed. The default value is **False** . Read/write **Boolean** .|
|[EnableDrilldown](pivottable-enabledrilldown-property-excel.md)| **True** if drilldown is enabled. The default value is **True** . Read/write **Boolean** .|
|[EnableFieldDialog](pivottable-enablefielddialog-property-excel.md)| **True** if the **PivotTable Field** dialog box is available when the user double-clicks the PivotTable field. The default value is **True** . Read/write **Boolean** .|
|[EnableFieldList](pivottable-enablefieldlist-property-excel.md)| **False** to disable the ability to display the field list for the PivotTable. If the field list was already being displayed it disappears. The default value is **True** . Read/write **Boolean** .|
|[EnableWizard](pivottable-enablewizard-property-excel.md)| **True** if the **PivotTable Wizard** is available. The default value is **True** . Read/write **Boolean** .|
|[EnableWriteback](pivottable-enablewriteback-property-excel.md)| Returns or sets whether writing back to the data source is enabled for the specified PivotTable. The default value is **False** . Read/write.|
|[ErrorString](pivottable-errorstring-property-excel.md)|Returns or sets a  **String** value that represents the string displayed in cells that contain errors when the **[DisplayErrorString](pivottable-displayerrorstring-property-excel.md)** property is **True** .|
|[FieldListSortAscending](pivottable-fieldlistsortascending-property-excel.md)|Controls the sort order of fields in the PivotTable Field List. When this property is set to  **True** , the fields are sorted in ascending order. When it is set to **False** , the fields are sorted in data source order. Read/write.|
|[GrandTotalName](pivottable-grandtotalname-property-excel.md)|Returns or sets the text string label that is displayed in the grand total column or row heading in the specified PivotTable report. The default value is the string "Grand Total". Read/write  **String** .|
|||
|[Hidden](pivottable-hidden-property-excel.md)|Checks whether the PivotTable exists at the worksheet level.  **Boolean** . Read-only|
|[HiddenFields](pivottable-hiddenfields-property-excel.md)|Returns an object that represents either a single PivotTable field (a  **[PivotField](pivotfield-object-excel.md)** object) or a collection of all the fields (a **[PivotFields](pivotfields-object-excel.md)** object) that are currently not shown as row, column, page, or data fields. Read-only.|
|[InGridDropZones](pivottable-ingriddropzones-property-excel.md)|This property is used to toggle in-grid drop zones for a  **PivotTable** object. In some cases, it also affects the layout of the PivotTable. Read/write **Boolean** .|
|[InnerDetail](pivottable-innerdetail-property-excel.md)|Returns or sets the name of the field that will be shown as detail when the  **ShowDetail** property is **True** for the innermost row or column field. Read/write **String** .|
|[LayoutRowDefault](pivottable-layoutrowdefault-property-excel.md)|This property specifies the layout settings for PivotFields when they are added to the PivotTable for the first time. Read/write  **xlLayoutRowType** .|
|[Location](pivottable-location-property-excel.md)|Gets or sets a  **String** that represents the top-left cell in the body of the specified **[PivotTable](pivottable-object-excel.md)** . Read/write.|
|[ManualUpdate](pivottable-manualupdate-property-excel.md)| **True** if the PivotTable report is recalculated only at the user's request. The default value is **False** . Read/write **Boolean** .|
|[MDX](pivottable-mdx-property-excel.md)|Returns a  **String** indicating the Multidimensional Expression (MDX) that would be sent to the provider to populate the current PivotTable view. Read-only.|
|[MergeLabels](pivottable-mergelabels-property-excel.md)| **True** if the specified PivotTable report's outer-row item, column item, subtotal, and grand total labels use merged cells. Read/write **Boolean** .|
|[Name](pivottable-name-property-excel.md)|Returns or sets a  **String** value representing the name of the object.|
|[NullString](pivottable-nullstring-property-excel.md)|Returns or sets the string displayed in cells that contain null values when the  **[DisplayNullString](pivottable-displaynullstring-property-excel.md)** property is **True** . The default value is an empty string (""). Read/write **String** .|
|[PageFieldOrder](pivottable-pagefieldorder-property-excel.md)|Returns or sets the order in which page fields are added to the PivotTable report's layout. Can be one of the following  **[XlOrder](xlorder-enumeration-excel.md)** constants: **xlDownThenOver** or **xlOverThenDown** . The default constant is **xlDownThenOver** . Read/write **Long** .|
|[PageFields](pivottable-pagefields-property-excel.md)|Returns an object that represents either a single PivotTable field (a  **[PivotField](pivotfield-object-excel.md)** object) or a collection of all the fields (a **[PivotFields](pivotfields-object-excel.md)** object) that are currently showing as page fields. Read-only.|
|[PageFieldStyle](pivottable-pagefieldstyle-property-excel.md)|Returns or sets the style used in the bound page field area. The default value is a null string (no style is applied by default). Read/write  **String** .|
|[PageFieldWrapCount](pivottable-pagefieldwrapcount-property-excel.md)|Returns or sets the number of page fields in each column or row in the PivotTable report. Read/write  **Long** .|
|[PageRange](pivottable-pagerange-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents the range that contains the page area in the PivotTable report. Read-only.|
|[PageRangeCells](pivottable-pagerangecells-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents only the cells in the specified PivotTable report that contain the page fields and item drop-down lists.|
|[Parent](pivottable-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[PivotChart](pivottable-pivotchart-property-excel.md)|Returns a [Shape Object (Excel)](shape-object-excel.md) object that represents the standalone PivotChart for the specified hidden PivotTable report. Read-only.|
|[PivotColumnAxis](pivottable-pivotcolumnaxis-property-excel.md)|Returns a  **PivotAxis** object representing the entire column axis. Read-only **PivotAxis** .|
|[PivotFormulas](pivottable-pivotformulas-property-excel.md)|Returns a  **[PivotFormulas](pivotformulas-object-excel.md)** object that represents the collection of formulas for the specified PivotTable report. Read-only.|
|[PivotRowAxis](pivottable-pivotrowaxis-property-excel.md)|Returns a  **PivotAxis** object representing the entire row axis. Read-only **PivotAxis** .|
|[PivotSelection](pivottable-pivotselection-property-excel.md)|Returns or sets the PivotTable selection in standard PivotTable report selection format. Read/write  **String** .|
|[PivotSelectionStandard](pivottable-pivotselectionstandard-property-excel.md)|Returns or sets a  **String** indicating the PivotTable selection in standard PivotTable report format using English (United States) settings. Read/write.|
|[PreserveFormatting](pivottable-preserveformatting-property-excel.md)| **True** if formatting is preserved when the report is refreshed or recalculated by operations such as pivoting, sorting, or changing page field items.For query tables, this property is **True** if any formatting common to the first five rows of data are applied to new rows of data in the query table. Unused cells aren't formatted. The property is **False** if the last AutoFormat applied to the query table is applied to new rows of data. The default value is **True** .|
|[PrintDrillIndicators](pivottable-printdrillindicators-property-excel.md)|Specifies whether or not drill indicators are printed with the PivotTable. Read/write  **Boolean** .|
|[PrintTitles](pivottable-printtitles-property-excel.md)| **True** if the print titles for the worksheet are set based on the PivotTable report. **False** if the print titles for the worksheet are used. The default value is **False** . Read/write **Boolean** .|
|[RefreshDate](pivottable-refreshdate-property-excel.md)|Returns the date on which the PivotTable report was last refreshed. Read-only  **Date** .|
|[RefreshName](pivottable-refreshname-property-excel.md)|Returns the name of the person who last refreshed the PivotTable report data. Read-only  **String** .|
|[RepeatItemsOnEachPrintedPage](pivottable-repeatitemsoneachprintedpage-property-excel.md)| **True** if row, column, and item labels appear on the first row of each page when the specified PivotTable report is printed. **False** if labels are printed only on the first page. The default value is **True** . Read/write **Boolean** .|
|[RowFields](pivottable-rowfields-property-excel.md)|Returns an object that represents either a single field in a PivotTable report (a  **[PivotField](pivotfield-object-excel.md)** object) or a collection of all the fields (a **[PivotFields](pivotfields-object-excel.md)** object) that are currently showing as row fields. Read-only.|
|[RowGrand](pivottable-rowgrand-property-excel.md)| **True** if the PivotTable report shows grand totals for rows. Read/write **Boolean** .|
|[RowRange](pivottable-rowrange-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents the range including the row area on the PivotTable report. Read-only.|
|[SaveData](pivottable-savedata-property-excel.md)| **True** if data for the PivotTable report is saved with the workbook. **False** if only the report definition is saved. Read/write **Boolean** .|
|[SelectionMode](pivottable-selectionmode-property-excel.md)|Returns or sets the PivotTable report structured selection mode. Read/write  **[XlPTSelectionMode](xlptselectionmode-enumeration-excel.md)** .|
|[ShowDrillIndicators](pivottable-showdrillindicators-property-excel.md)|The  **ShowDrillIndicators** property is used for toggling the display of drill indicators in the PivotTable. Read/write **Boolean** .|
|[ShowPageMultipleItemLabel](pivottable-showpagemultipleitemlabel-property-excel.md)|When set to  **True** (default), "(Multiple Items)" will appear in the PivotTable cell on the worksheet whenever items are hidden and an aggregate of non-hidden items is shown in the PivotTable view. Read/write **Boolean** .|
|[ShowTableStyleColumnHeaders](pivottable-showtablestylecolumnheaders-property-excel.md)|The  **ShowTableStyleColumnHeaders** property is set to **True** if the coulmn headers should be displayed in the PivotTable. Read/write **Boolean** .|
|[ShowTableStyleColumnStripes](pivottable-showtablestylecolumnstripes-property-excel.md)|The  **ShowTableStyleColumnStripes** property displays banded columns in which even columns are formatted differently from odd columns. This makes PivotTables easier to read. Read/write **Boolean** .|
|||
|[ShowTableStyleRowHeaders](pivottable-showtablestylerowheaders-property-excel.md)|The  **ShowTableStyleRowHeaders** property is set to **True** if the row headers should be displayed in the PivotTable. Read/write **Boolean** .|
|[ShowTableStyleRowStripes](pivottable-showtablestylerowstripes-property-excel.md)|The  **ShowTableStyleRowStripes** property displays banded rows in which even rows are formatted differently from odd rows. This makes PivotTables easier to read. Read/write **Boolean** .|
|[ShowValuesRow](pivottable-showvaluesrow-property-excel.md)|Returns or sets whether the values row is displayed. Read/write|
|[Slicers](pivottable-slicers-property-excel.md)|Returns the  **[Slicers](slicers-object-excel.md)** collection for the specified PivotTable. Read-only|
|[SmallGrid](pivottable-smallgrid-property-excel.md)| **True** if Microsoft Excel uses a grid that's two cells wide and two cells deep for a newly created PivotTable report. **False** if Excel uses a blank stencil outline. Read/write **Boolean** .|
|[SortUsingCustomLists](pivottable-sortusingcustomlists-property-excel.md)|The  **SortUsingCustomLists** property controls whether custom lists are used for sorting items of fields, both initially when the PivotField is initialized and the PivotItems are ordered by their captions; and later when the user applies a sort. Read/write **Boolean** .|
|[SourceData](pivottable-sourcedata-property-excel.md)|Returns the data source for the PivotTable report, as shown in the following table. Read-write  **Variant** .|
|[SubtotalHiddenPageItems](pivottable-subtotalhiddenpageitems-property-excel.md)| **True** if hidden page field items in the PivotTable report are included in row and column subtotals, block totals, and grand totals. The default value is **False** . Read/write **Boolean** .|
|[Summary](pivottable-summary-property-excel.md)|Returns or sets the description associated with the alternative text string for the specified PivotTable. Read/write|
|[TableRange1](pivottable-tablerange1-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents the range containing the entire PivotTable report, but doesn't include page fields. Read-only.|
|[TableRange2](pivottable-tablerange2-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents the range containing the entire PivotTable report, including page fields. Read-only.|
|[TableStyle2](pivottable-tablestyle2-property-excel.md)|The  **TableStyle2** property specifies the PivotTable style currently applied to the PivotTable. Read/write.|
|[Tag](pivottable-tag-property-excel.md)|Returns or sets a string saved with the PivotTable report. Read/write  **String** .|
|[TotalsAnnotation](pivottable-totalsannotation-property-excel.md)| **True** if an asterisk (*) is displayed next to each subtotal and grand total value in the specified PivotTable report if the report is based on an OLAP data source. The default value is **True** . Read/write **Boolean** .|
|[VacatedStyle](pivottable-vacatedstyle-property-excel.md)|Returns or sets the style applied to cells vacated when the PivotTable report is refreshed. The default value is a null string (no style is applied by default). Read/write  **String** .|
|[Value](pivottable-value-property-excel.md)|Returns or sets a  **String** value that represents the name of the PivotTable report.|
|[Version](pivottable-version-property-excel.md)|Returns a  **[XlPivotTableVersionList](xlpivottableversionlist-enumeration-excel.md)** value that represents the Microsoft Excel version number.|
|[ViewCalculatedMembers](pivottable-viewcalculatedmembers-property-excel.md)|When set to  **True** (default), calculated members for Online Analytical Processing (OLAP) PivotTables can be viewed. Read/write **Boolean** .|
|[VisibleFields](pivottable-visiblefields-property-excel.md)|Returns an object that represents either a single field in a PivotTable report (a  **[PivotField](pivotfield-object-excel.md)** object) or a collection of all the visible fields (a **[PivotFields](pivotfields-object-excel.md)** object). Visible fields are shown as row, column, page or data fields. Read-only.|
|[VisualTotals](pivottable-visualtotals-property-excel.md)| **True** (default) to enable Online Analytical Processing (OLAP) PivotTables to retotal after an item has been hidden from view. Read/write **Boolean** .|
|[VisualTotalsForSets](pivottable-visualtotalsforsets-property-excel.md)|Returns or sets whether to include filtered items in the totals of named sets for the specified PivotTable. Read/write|

