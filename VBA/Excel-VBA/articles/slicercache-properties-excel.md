---
title: SlicerCache Properties (Excel)
ms.prod: EXCEL
ms.assetid: dafc4a75-a25f-4c9e-894f-c3778a2a7c27
---


# SlicerCache Properties (Excel)

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](slicercache-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object. Read-only.|
|[Creator](slicercache-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[CrossFilterType](slicercache-crossfiltertype-property-excel.md)|Returns or sets whether a slicer is participating in cross filtering with other slicers that share the same slicer cache, and how cross filtering is displayed. Read/write|
|[FilterCleared](slicercache-filtercleared-property-excel.md)|Returns whether the slicer or timeline filter state is cleared.  **Boolean** . Read-only|
|[Index](slicercache-index-property-excel.md)| Returns the index of the specified **[SlicerCache](slicercache-object-excel.md)** object in the **[SlicerCaches](slicercaches-object-excel.md)** collection. Read-only|
|[List](slicercache-list-property-excel.md)| **True** if the slicer cache is for a slicer on a table. **False** otherwise. **Boolean** Read-only|
|[ListObject](slicercache-listobject-property-excel.md)|Returns a  **ListObject** object for the[QueryTable Object (Excel)](querytable-object-excel.md) object. Read-only|
|[Name](slicercache-name-property-excel.md)|Returns or sets the name of the slicer cache.|
|[OLAP](slicercache-olap-property-excel.md)|Returns whether the slicer associated with the specified slicer cache is based on an OLAP data source. Read-only|
|[Parent](slicercache-parent-property-excel.md)|Returns the parent  **[SlicerCaches](slicercaches-object-excel.md)** object for the specified **SlicerCache** object. Read-only.|
|[PivotTables](slicercache-pivottables-property-excel.md)|Returns a  **[SlicerPivotTables](slicerpivottables-object-excel.md)** collection that contains information about the PivotTables the slicer cache is currently filtering. Read-only|
|[RequireManualUpdate](slicercache-requiremanualupdate-property-excel.md)| **True** when manual updates of the slicer cache required. **Boolean** Read/Write|
|[ShowAllItems](slicercache-showallitems-property-excel.md)|Returns or sets whether slicers connected to the specified slicer cache display items that have been deleted from in the corresponding PivotCache. Read/write|
|[SlicerCacheLevels](slicercache-slicercachelevels-property-excel.md)|Returns the collection of  **[SlicerCacheLevel](slicercachelevel-object-excel.md)** objects that represent the levels of an OLAP hierarchy on which the specified slicer cache is based. Read-only|
|[SlicerCacheType](slicercache-slicercachetype-property-excel.md)|Returns the type of the slicer cache - slicer or timeline. Read-only|
|[SlicerItems](slicercache-sliceritems-property-excel.md)|Returns a  **[SlicerItems](sliceritems-object-excel.md)** collection that contains the collection of all items in the slicer cache. Read-only|
|[Slicers](slicercache-slicers-property-excel.md)|Returns a  **[Slicers](slicers-object-excel.md)** collection that contains the collection of **[Slicer](slicer-object-excel.md)** objects associated with the specified **[SlicerCache](slicercache-object-excel.md)** . Read-only|
|[SortItems](slicercache-sortitems-property-excel.md)|Returns or sets the sort order of the items in the slicer. Read/write  **[XlSlicerSort](xlslicersort-enumeration-excel.md)** .|
|[SortUsingCustomLists](slicercache-sortusingcustomlists-property-excel.md)|Returns or sets whether items in the specified slicer cache will be sorted by the custom lists. Read/write|
|[SourceName](slicercache-sourcename-property-excel.md)|Returns the name of the data source the slicer is connected to. Read-only|
|[SourceType](slicercache-sourcetype-property-excel.md)|Returns the kind of data source the slicer is connected to. Read-only|
|[TimelineState](slicercache-timelinestate-property-excel.md)|The timeline-specific state of the  **SlicerCache** object. Read-only|
|[VisibleSlicerItems](slicercache-visiblesliceritems-property-excel.md)|Returns a  **[SlicerItems](sliceritems-object-excel.md)** collection that contains the collection of all the visible items in the specified slicer cache. Read-only|
|[VisibleSlicerItemsList](slicercache-visiblesliceritemslist-property-excel.md)|Returns or sets the list of MDX unique names for members at all levels of the hierarchy where manual filtering is applied. Read/write|
|[WorkbookConnection](slicercache-workbookconnection-property-excel.md)|Gets or sets the  **[WorkbookConnection](workbookconnection-object-excel.md)** object that represents the data connection used by the specified slicer. Read/Write.|

