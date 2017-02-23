---
title: ValueChange Members (Excel)
ms.prod: EXCEL
ms.assetid: cd467d92-dee0-d049-0457-ec85ef74adf8
---


# ValueChange Members (Excel)
Represents a value that has been changed in a PivotTable report that is based on an OLAP data source.

Represents a value that has been changed in a PivotTable report that is based on an OLAP data source.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Delete](valuechange-delete-method-excel.md)|Deletes the specified  **[ValueChange](valuechange-object-excel.md)** object from the **[PivotTableChangeList](pivottablechangelist-object-excel.md)** collection.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[AllocationMethod](valuechange-allocationmethod-property-excel.md)|Returns what method to use to allocate this value when performing what-if analysis. Read-only|
|[AllocationValue](valuechange-allocationvalue-property-excel.md)|Returns what value to allocate when performing what-if analysis. Read-only|
|[AllocationWeightExpression](valuechange-allocationweightexpression-property-excel.md)|Returns the MDX weight expression to use for this value when performing what-if analysis. Read-only|
|[Application](valuechange-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object. Read-only.|
|[Creator](valuechange-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[Order](valuechange-order-property-excel.md)|Returns a value that indicates the order in which this change was performed relative to other changes in the  **[PivotTableChangeList](pivottablechangelist-object-excel.md)** collection. Read-only|
|[Parent](valuechange-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[PivotCell](valuechange-pivotcell-property-excel.md)|Returns a  **[PivotCell](pivotcell-object-excel.md)** object that represents the cell (tuple) that was changed. Read-only|
|[Tuple](valuechange-tuple-property-excel.md)|Returns the MDX tuple of the value was changed in the OLAP data source. Read-only|
|[Value](valuechange-value-property-excel.md)|Returns the value that the user entered in the cell or that the formula in the cell was evaluated to when  **UPDATE CUBE** statement was last run against the OLAP data source. Read-only|
|[VisibleInPivotTable](valuechange-visibleinpivottable-property-excel.md)|Returns whether the cell (tuple) is currently visible in the PivotTable report. Read-only|

