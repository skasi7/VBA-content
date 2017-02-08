---
title: TimelineState Members (Excel)
ms.prod: EXCEL
ms.assetid: 6c21dcbb-b0a6-0f24-27f6-6aefafc5f6ec
---


# TimelineState Members (Excel)
The timeline-specific state of a [SlicerCache Object (Excel)](slicercache-object-excel.md) object.

The timeline-specific state of a [SlicerCache Object (Excel)](slicercache-object-excel.md) object.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[SetFilterDateRange](timelinestate-setfilterdaterange-method-excel.md)|Sets the Timeline's filter.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](timelinestate-application-property-excel.md)|Returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. Read-only.|
|[Creator](timelinestate-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[EndDate](timelinestate-enddate-property-excel.md)|Returns the end of the filtering date range (equals to [TimelineState.StartDate Property (Excel)](timelinestate-startdate-property-excel.md) if range is a single day). **Variant** Read-only|
|[FilterType](timelinestate-filtertype-property-excel.md)|Returns the type of the date filter. [XlPivotFilterType Enumeration (Excel)](xlpivotfiltertype-enumeration-excel.md) Read-only|
|[FilterValue1](timelinestate-filtervalue1-property-excel.md)|Returns the 1st value associated with the date filter (semantics vary by filter type).  **Variant** Read-only|
|[FilterValue2](timelinestate-filtervalue2-property-excel.md)|Returns the 2nd value associated with the date filter (semantics vary by filter type).  **Variant** Read-only|
|[Parent](timelinestate-parent-property-excel.md)|Returns an  **Object** that represents the parent object of the specified[TimelineState Object (Excel)](timelinestate-object-excel.md) object. Read-only.|
|[SingleRangeFilterState](timelinestate-singlerangefilterstate-property-excel.md)| **True** when the filtering state is a contiguous date range; **False** otherwise. **Boolean** Read-only|
|[StartDate](timelinestate-startdate-property-excel.md)|Returns the start of the filtering date range.  **Variant** Read-only|

