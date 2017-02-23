---
title: PivotFilter Members (Excel)
ms.prod: EXCEL
ms.assetid: a1be2481-9d14-cc49-8a1b-187048f0d179
---


# PivotFilter Members (Excel)
A PivotFilter is applied to a  **PivotField** object.

A PivotFilter is applied to a  **PivotField** object.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Delete](pivotfilter-delete-method-excel.md)|Deletes the filter and removes it from the filter collections of the PivotField and the PivotTable.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Active](pivotfilter-active-property-excel.md)|Returns whether the specified PivotFilter is active. Read-only  **Boolean** .|
|[Application](pivotfilter-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object. Read-only.|
|[Creator](pivotfilter-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[DataCubeField](pivotfilter-datacubefield-property-excel.md)|This property is applicable only to OLAP PivotTables and provides the  **Value** field (PivotField in the Values area) being filtered by for a value filter. Read/write **CubeField** .|
|[DataField](pivotfilter-datafield-property-excel.md)|This property is applicable only to non-OLAP PivotTables and provides the  **Value** field (PivotField in the Values area) being filtered by for a value filter. Read/write **PivotField** .|
|[Description](pivotfilter-description-property-excel.md)|Provides an optional description for the  **PivotFilter** object. Read-only **String** .|
|[FilterType](pivotfilter-filtertype-property-excel.md)|Specifies the type of filter to be applied. Read-only  **xlPivotFilterType** .|
|[IsMemberPropertyFilter](pivotfilter-ismemberpropertyfilter-property-excel.md)|Specifies whether the label filter is based on the PivotItem captions of a member property of the field or on the PivotItem captions of the PivotField itself. Read-only  **Boolean** .|
|[MemberPropertyField](pivotfilter-memberpropertyfield-property-excel.md)|This property specifies the member property PivotField on which the label filter is based. Read/write  **PivotField** .|
|[Name](pivotfilter-name-property-excel.md)|This property provides the option of naming filters for reference. You cannot rely on the index value for accurate reference because this value can change.|
|[Order](pivotfilter-order-property-excel.md)|Specifies the evaluation order of the filter among all Value filters applied to the entire PivotTable. Read/write  **Integer** .|
|[Parent](pivotfilter-parent-property-excel.md)|Returns the parent object for the specified  **PivotFilter** object. Read-only.|
|[PivotField](pivotfilter-pivotfield-property-excel.md)|Specifies the PivotField to which the filter is applied. Read-only.|
|[Value1](pivotfilter-value1-property-excel.md)|This property is a user-supplied parameter to define a filter for a PivotField. Read/write  **Variant** .|
|[Value2](pivotfilter-value2-property-excel.md)|This property is a user-supplied parameter to define a filter for a PivotField. Read/write  **Variant** .|
|[WholeDayFilter](pivotfilter-wholedayfilter-property-excel.md)|Sets or gets the filtering semantics for date filters.  **Boolean** . Read/Write|

