---
title: CalculatedMember Properties (Excel)
ms.prod: EXCEL
ms.assetid: dfaea829-b425-462b-98e4-0fc3124310d8
---


# CalculatedMember Properties (Excel)

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](calculatedmember-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[Creator](calculatedmember-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[DisplayFolder](calculatedmember-displayfolder-property-excel.md)|Returns the display folder name for a named set. Read-only|
|[Dynamic](calculatedmember-dynamic-property-excel.md)|Returns whether the specified named set is recalculated with every update. Read-only|
|[FlattenHierarchies](calculatedmember-flattenhierarchies-property-excel.md)|Returns or sets whether items from all levels of the hierarchy of the specified named set are displayed in the same field of a PivotTable report based on an OLAP cube.  **Boolean** Read/write|
|[Formula](calculatedmember-formula-property-excel.md)|Returns a  **String** value that represents the member's formula in multidimensional expressions (MDX) syntax.|
|[HierarchizeDistinct](calculatedmember-hierarchizedistinct-property-excel.md)|Returns or sets whether to order and remove duplicates when displaying the hierarchy of the specified named set in a PivotTable report based on an OLAP cube. Read/write|
|[IsValid](calculatedmember-isvalid-property-excel.md)|Returns a Boolean that indicates whether the specified calculated member has been successfully instantiated with the OLAP provider during the current session.|
|[MeasureGroup](calculatedmember-measuregroup-property-excel.md)|Returns the associated measure group.  **String** Read-only|
|[Name](calculatedmember-name-property-excel.md)|Returns a  **String** value that represents the name of the object.|
|[NumberFormat](calculatedmember-numberformat-property-excel.md)|Returns a  **[XlCalcMemNumberFormatType](xlcalcmemnumberformattype-enumeration-excel.md)** value that represents the number format of the calculated member. The default value is **xlNumberFormatTypeDefault** . Read-only.|
|[Parent](calculatedmember-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[ParentHierarchy](calculatedmember-parenthierarchy-property-excel.md)| Returns the name of the current parent hierarchy from the hierarchies that are available on the cube. **String** Read-only|
|[ParentMember](calculatedmember-parentmember-property-excel.md)|Returns the name of the parent member for the parent hierarchy.  **String** Read-only|
|[SolveOrder](calculatedmember-solveorder-property-excel.md)|Returns a  **Long** specifying the value of the calculated member's solve order MDX (Mulitdimensional Expression) argument. The default value is zero. Read-only.|
|[SourceName](calculatedmember-sourcename-property-excel.md)|Returns a  **String** value that represents the specified object's name as it appears in the original source data for the specified PivotTable report.|
|[Type](calculatedmember-type-property-excel.md)|Returns a  **[XlCalculatedMemberType](xlcalculatedmembertype-enumeration-excel.md)** value that represents the calculated member type.|

