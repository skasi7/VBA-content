---
title: PivotField Members (Excel)
ms.prod: EXCEL
ms.assetid: 4a6ea12a-072c-a386-c855-7bf5f6eadd46
---


# PivotField Members (Excel)
Represents a field in a PivotTable report.

Represents a field in a PivotTable report.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[AddPageItem](pivotfield-addpageitem-method-excel.md)|Adds an additional item to a multiple item page field.|
|[AutoShow](pivotfield-autoshow-method-excel.md)|Displays the number of top or bottom items for a row, page, or column field in the specified PivotTable report.|
|[AutoSort](pivotfield-autosort-method-excel.md)|Establishes automatic field-sorting rules for PivotTable reports.|
|[CalculatedItems](pivotfield-calculateditems-method-excel.md)|Returns a  **[CalculatedItems](calculateditems-object-excel.md)** collection that represents all the calculated items in the specified PivotTable report. Read-only.|
|[ClearAllFilters](pivotfield-clearallfilters-method-excel.md)|Calling this method deletes all filters currently applied to the PivotField. This includes deleting all filters from the  **PivotFilters** collection of the PivotField and removing any manual filtering applied to the PivotField as well. If the PivotField is in the Report Filter area, the item selected will be set to the default item.|
|[ClearLabelFilters](pivotfield-clearlabelfilters-method-excel.md)|This method deletes all label filters or all date filters in the  **PivotFilters** collection of the PivotField.|
|[ClearManualFilter](pivotfield-clearmanualfilter-method-excel.md)|Provides an easy way to set the  **Visible** property to **True** for all items of a PivotField in PivotTables, and to empty the **HiddenItemsList** and **VisibleItemsList** collections in OLAP PivotTables.|
|[ClearValueFilters](pivotfield-clearvaluefilters-method-excel.md)|Calling this method deletes all value filters in the  **PivotFilters** collection of the PivotField.|
|[Delete](pivotfield-delete-method-excel.md)|Deletes the object.|
|[DrillTo](pivotfield-drillto-method-excel.md)|The  **DrillTo** method supports drilling to a specified PivotField from another PivotField.|
|[PivotItems](pivotfield-pivotitems-method-excel.md)|Returns an object that represents either a single PivotTable item (a  **[PivotItem](pivotitem-object-excel.md)** object) or a collection of all the visible and hidden items (a **[PivotItems](pivotitems-object-excel.md)** object) in the specified field. Read-only.|
|[AutoGroup](pivotfield-autogroup-method-excel.md)|Automatically groups the pivot fields in a pivot table.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[AllItemsVisible](pivotfield-allitemsvisible-property-excel.md)|Used to retrieve a Boolean value that indicates whether or not any manual filtering is applied to the PivotField. Read-only.|
|[Application](pivotfield-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[AutoShowCount](pivotfield-autoshowcount-property-excel.md)|Returns the number of top or bottom items that are automatically shown in the specified PivotTable field. Read-only  **Long** .|
|[AutoShowField](pivotfield-autoshowfield-property-excel.md)|Returns the name of the data field used to determine the top or bottom items that are automatically shown in the specified PivotTable field. Read-only  **String** .|
|[AutoShowRange](pivotfield-autoshowrange-property-excel.md)|Returns  **xlTop** if the top items are shown automatically in the specified PivotTable field; returns **xlBottom** if the bottom items are shown. Read-only **Long** .|
|[AutoShowType](pivotfield-autoshowtype-property-excel.md)|Returns  **xlAutomatic** if **AutoShow** is enabled for the specified PivotTable field; returns **xlManual** if **AutoShow** is disabled. Read-only **Long** .|
|[AutoSortCustomSubtotal](pivotfield-autosortcustomsubtotal-property-excel.md)|Returns the name of the custom subtotal used to sort the specified PivotTable field automatically. Read-only.|
|[AutoSortField](pivotfield-autosortfield-property-excel.md)|Returns the name of the data field used to sort the specified PivotTable field automatically. Read-only  **String** .|
|[AutoSortOrder](pivotfield-autosortorder-property-excel.md)|Returns the order used to sort the specified PivotTable field automatically. Can be one of the constants of  **[XlSortOrder](xlsortorder-enumeration-excel.md)** . Read-only **Long** .|
|[AutoSortPivotLine](pivotfield-autosortpivotline-property-excel.md)|Returns the name of the PivotLine used to sort the specified PivotTable field automatically. Read-only.|
|[BaseField](pivotfield-basefield-property-excel.md)|Returns or sets the base field for a custom calculation. This property is valid only for data fields. Read/write  **Variant** .|
|[BaseItem](pivotfield-baseitem-property-excel.md)|Returns or sets the item in the base field for a custom calculation. Valid only for data fields. Read/write  **Variant** .|
|[Calculation](pivotfield-calculation-property-excel.md)|Returns or sets a  **[XlPivotFieldCalculation](xlpivotfieldcalculation-enumeration-excel.md)** value that represents the type of calculation performed by the specified field. This property is valid only for data fields.|
|[Caption](pivotfield-caption-property-excel.md)|Returns a  **String** value that represents the label text for the pivot field.|
|[ChildField](pivotfield-childfield-property-excel.md)|Returns a  **[PivotField](pivotfield-object-excel.md)** object that represents the child field for the specified field (if the field is grouped and has a child field). Read-only.|
|[ChildItems](pivotfield-childitems-property-excel.md)|Returns an object that represents either a single PivotTable item (a  **[PivotItem](pivotitem-object-excel.md)** object) or a collection of all the items (a **[PivotItems](pivotitems-object-excel.md)** object) that are group children in the specified field, or children of the specified item. Read-only.|
|[Creator](pivotfield-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[CubeField](pivotfield-cubefield-property-excel.md)|Returns the  **[CubeField](cubefield-object-excel.md)** object from which the specified PivotTable field is descended. Read-only.|
|[CurrentPage](pivotfield-currentpage-property-excel.md)|Returns or sets the current page showing for the page field (valid only for page fields). Read/write  **[PivotItem](pivotitem-object-excel.md)** .|
|[CurrentPageList](pivotfield-currentpagelist-property-excel.md)|Returns or sets an array of strings corresponding to the list of items included in a multiple-item page field of a PivotTable report. Read/write  **Variant** .|
|[CurrentPageName](pivotfield-currentpagename-property-excel.md)|Returns or sets the currently displayed page of the specified PivotTable report. The name of the page appears in the page field. Note that this property works only if the currently displayed page already exists. Read/write  **String** .|
|[DatabaseSort](pivotfield-databasesort-property-excel.md)|When set to  **True** , manual repositioning of items in a PivotTable field is allowed. Returns **True** , if the field has no manually positioned items. Read/write **Boolean** .|
|[DataRange](pivotfield-datarange-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object as shown in the following table. Read-only.|
|[DataType](pivotfield-datatype-property-excel.md)|Returns a  **[XlPivotFieldDataType](xlpivotfielddatatype-enumeration-excel.md)** value that represents the type of data in the PivotTable field.|
|[DisplayAsCaption](pivotfield-displayascaption-property-excel.md)|This property is used to display member properties of PivotFields as captions. Read-only.|
|[DisplayAsTooltip](pivotfield-displayastooltip-property-excel.md)|This property is used to specify whether or not a specific member property PivotField is displayed in tooltips. Read/write  **Boolean** .|
|[DisplayInReport](pivotfield-displayinreport-property-excel.md)|This property is used to specify whether the specified member property PivotField is displayed in the PivotTable or not. Read/write  **Boolean** .|
|[DragToColumn](pivotfield-dragtocolumn-property-excel.md)| **True** if the specified field can be dragged to the column position. The default value is **True** . Read/write **Boolean** .|
|[DragToData](pivotfield-dragtodata-property-excel.md)| **True** if the specified field can be dragged to the data position. The default value is **True** . Read/write **Boolean**|
|[DragToHide](pivotfield-dragtohide-property-excel.md)| **True** if the field can be hidden by being dragged off the PivotTable report. The default value is **True** . Read/write **Boolean** .|
|[DragToPage](pivotfield-dragtopage-property-excel.md)| **True** if the field can be dragged to the page position. The default value is **True** . Read/write **Boolean** .|
|[DragToRow](pivotfield-dragtorow-property-excel.md)| **True** if the field can be dragged to the row position. The default value is **True** . Read/write **Boolean** .|
|[DrilledDown](pivotfield-drilleddown-property-excel.md)| **True** if the flag for the specified PivotTable field or PivotTable item is set to "drilled" (expanded, or visible). Read/write **Boolean** .|
|[EnableItemSelection](pivotfield-enableitemselection-property-excel.md)|When set to  **False** , disables the ability to use the field dropdown in the user interface. The default value is **True** . Read/write **Boolean** .|
|[EnableMultiplePageItems](pivotfield-enablemultiplepageitems-property-excel.md)|Used for specifying whether or not check boxes are present in the filter drop-down list for fields in the page area. Read/write  **Boolean** .|
|[Formula](pivotfield-formula-property-excel.md)|Returns or sets a  **String** value that represents the object's formula in A1-style notation and in the language of the macro.|
|[Function](pivotfield-function-property-excel.md)|Returns or sets the function used to summarize the PivotTable field (data fields only). Read/write  **[XlConsolidationFunction](xlconsolidationfunction-enumeration-excel.md)** .|
|[GroupLevel](pivotfield-grouplevel-property-excel.md)|Returns the placement of the specified field within a group of fields (if the field is a member of a grouped set of fields). Read-only.|
|[Hidden](pivotfield-hidden-property-excel.md)|This property is used to hide the individual levels of an OLAP hierarchy. Read/write  **Boolean** .|
|[HiddenItems](pivotfield-hiddenitems-property-excel.md)|Returns an object that represents either a single hidden PivotTable item (a  **[PivotItem](pivotitem-object-excel.md)** object) or a collection of all the hidden items (a **[PivotItems](pivotitems-object-excel.md)** object) in the specified field. Read-only.|
|[HiddenItemsList](pivotfield-hiddenitemslist-property-excel.md)|Returns or sets a  **Variant** specifying an array of strings that are hidden items for a PivotTable field. Read/write.|
|[IncludeNewItemsInFilter](pivotfield-includenewitemsinfilter-property-excel.md)|Allows developers to specify whether excluded or included items should be tracked when manual filtering is applied to the PivotField. Read/write  **Boolean** .|
|[IsCalculated](pivotfield-iscalculated-property-excel.md)| **True** if the PivotTable field is a calculated field or item. Read-only **Boolean** .|
|[IsMemberProperty](pivotfield-ismemberproperty-property-excel.md)|Returns  **True** when the PivotField contains member properties. Read-only **Boolean** .|
|[LabelRange](pivotfield-labelrange-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents the cell (or cells) that contain the field label. Read-only|
|[LayoutBlankLine](pivotfield-layoutblankline-property-excel.md)| **True** if a blank row is inserted after the specified row field in a PivotTable report. The default value is **False** . Read/write **Boolean** .|
|[LayoutCompactRow](pivotfield-layoutcompactrow-property-excel.md)|Specifies whether or not a PivotField is compacted (items of multiple PivotFields are displayed in a single column) when rows are selected. Read/write  **Boolean** .|
|[LayoutForm](pivotfield-layoutform-property-excel.md)|Returns or sets the way the specified PivotTable items appear—in table format or in outline format. Read/write  **[XlLayoutFormType](xllayoutformtype-enumeration-excel.md)** .|
|[LayoutPageBreak](pivotfield-layoutpagebreak-property-excel.md)| **True** if a page break is inserted after each field. The default value is **False** . Read/write **Boolean** .|
|[LayoutSubtotalLocation](pivotfield-layoutsubtotallocation-property-excel.md)|Returns or sets the position of the PivotTable field subtotals in relation to (either above or below) the specified field. Read/write  **[XlSubtototalLocationType](xlsubtototallocationtype-enumeration-excel.md)** .|
|[MemberPropertyCaption](pivotfield-memberpropertycaption-property-excel.md)|Setting the  **MemberPropertyCaption** property controls which member property is used as caption for a given level. Read/write **Boolean** .|
|[MemoryUsed](pivotfield-memoryused-property-excel.md)|Returns the amount of memory currently being used by the object, in bytes. Read-only  **Long** .|
|[Name](pivotfield-name-property-excel.md)|Returns or sets a  **String** value representing the name of the object.|
|[NumberFormat](pivotfield-numberformat-property-excel.md)|Returns or sets a  **String** value that represents the format code for the object.|
|[Orientation](pivotfield-orientation-property-excel.md)|Returns or sets a  **[XlPivotFieldOrientation](xlpivotfieldorientation-enumeration-excel.md)** value that represents the location of the field in the specified PivotTable report.|
|[Parent](pivotfield-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[ParentField](pivotfield-parentfield-property-excel.md)|Returns a  **[PivotField](pivotfield-object-excel.md)** object that represents the PivotTable field that's the group parent of the specified object. The field must be grouped and must have a parent field. Read-only.|
|[ParentItems](pivotfield-parentitems-property-excel.md)|Returns an object that represents either a single PivotTable item (a  **[PivotItem](pivotitem-object-excel.md)** object) or a collection of all the items (a **[PivotItems](pivotitems-object-excel.md)** object) that are group parents in the specified field. The specified field must be a group parent of another field. Read-only.|
|[PivotFilters](pivotfield-pivotfilters-property-excel.md)|Returns or sets the PivotFilters for the specified  **PivotField** object. Read-only.|
|[Position](pivotfield-position-property-excel.md)|Returns or sets a Variant value that represents the position of the field (first, second, third, and so on) among all the fields in its orientation (Rows, Columns, Pages, Data).|
|[PropertyOrder](pivotfield-propertyorder-property-excel.md)|Valid only for PivotTable fields that are member property fields. Returns a  **Long** indicating the display position of the member property within the cube field to which it belongs. Read/write.|
|[PropertyParentField](pivotfield-propertyparentfield-property-excel.md)|Returns a  **PivotField** object representing the field to which the properties in this field pertain.|
|[RepeatLabels](pivotfield-repeatlabels-property-excel.md)|Returns or sets whether item labels are repeated in the PivotTable for the specified PivotField. Read/write|
|[ServerBased](pivotfield-serverbased-property-excel.md)| **True** if the data source for the specified PivotTable report is external and only the items matching the page field selection are retrieved. Read/write **Boolean** .|
|[ShowAllItems](pivotfield-showallitems-property-excel.md)| **True** if all items in the PivotTable report are displayed, even if they don't contain summary data. The default value is **False** . Read/write **Boolean** .|
|[ShowDetail](pivotfield-showdetail-property-excel.md)|Gets or sets whether the specified  **[PivotField](pivotfield-object-excel.md)** is showing detail. Read/write **Boolean** .|
|[ShowingInAxis](pivotfield-showinginaxis-property-excel.md)|Indicates if the PivotField is currently visible in the PivotTable or not. Read-only.|
|[SourceCaption](pivotfield-sourcecaption-property-excel.md)|The  **SourceCaption** property is applicable only for OLAP PivotTables, and returns the original caption from the OLAP server for a PivotField. Read-only.|
|[SourceName](pivotfield-sourcename-property-excel.md)|Returns a  **String** value that represents the specified object's name as it appears in the original source data for the specified PivotTable report.|
|[StandardFormula](pivotfield-standardformula-property-excel.md)|Returns or sets a  **String** specifying formulas with standard English (United States) formatting. Read/write.|
|[SubtotalName](pivotfield-subtotalname-property-excel.md)|Returns or sets the text string label displayed in the subtotal column or row heading in the specified PivotTable report. The default value is the string "Subtotal". Read/write  **String** .|
|[Subtotals](pivotfield-subtotals-property-excel.md)|Returns or sets subtotals displayed with the specified field. Valid only for nondata fields. Read/write  **Variant** .|
|[TotalLevels](pivotfield-totallevels-property-excel.md)|Returns the total number of fields in the current field group. If the field isn't grouped, or if the data source is OLAP-based,  **TotalLevels** returns the value 1. Read-only **Long** .|
|[UseMemberPropertyAsCaption](pivotfield-usememberpropertyascaption-property-excel.md)|This property is used to control whether member property captions are used for PivotItem captions of the PivotField. Read/write  **Boolean** .|
|[Value](pivotfield-value-property-excel.md)|Returns or sets a  **String** value that represents the name of the specified field in the PivotTable report.|
|[VisibleItems](pivotfield-visibleitems-property-excel.md)|Returns an object that represents either a single visible PivotTable item (a  **[PivotItem](pivotitem-object-excel.md)** object) or a collection of all the visible items (a **[PivotItems](pivotitems-object-excel.md)** object) in the specified field. Read-only.|
|[VisibleItemsList](pivotfield-visibleitemslist-property-excel.md)|Returns or sets a  **Variant** specifying an array of strings that represent included items in a manual filter applied to a PivotField. Read/write.|

