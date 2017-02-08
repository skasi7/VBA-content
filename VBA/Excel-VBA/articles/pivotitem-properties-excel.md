---
title: PivotItem Properties (Excel)
ms.prod: EXCEL
ms.assetid: ffeaf68e-c366-448a-ab68-011acb60cd3f
---


# PivotItem Properties (Excel)

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](pivotitem-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[Caption](pivotitem-caption-property-excel.md)|Returns a  **String** value that represents the label text for the pivot item.|
|[ChildItems](pivotitem-childitems-property-excel.md)|Returns an object that represents either a single PivotTable item (a  **[PivotItem](pivotitem-object-excel.md)** object) or a collection of all the items (a **[PivotItems](pivotitems-object-excel.md)** object) that are group children in the specified field, or children of the specified item. Read-only.|
|[Creator](pivotitem-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[DataRange](pivotitem-datarange-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object as shown in the following table. Read-only.|
|[DrilledDown](pivotitem-drilleddown-property-excel.md)| **True** if the flag for the specified PivotTable field or PivotTable item is set to "drilled" (expanded, or visible). Read/write **Boolean** .|
|[Formula](pivotitem-formula-property-excel.md)|Returns or sets a  **String** value that represents the object's formula in A1-style notation and in the language of the macro.|
|[IsCalculated](pivotitem-iscalculated-property-excel.md)| **True** if the PivotTable item is a calculated field or item. Read-only **Boolean** .|
|[LabelRange](pivotitem-labelrange-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents all the cells in the PivotTable report that contain the item. Read-only.|
|[Name](pivotitem-name-property-excel.md)|Returns or sets a  **String** value representing the name of the object.|
|[Parent](pivotitem-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[ParentItem](pivotitem-parentitem-property-excel.md)|Returns a  **PivotItem** object that represents the parent PivotTable item in the parent **[PivotField](pivotfield-object-excel.md)** object (the field must be grouped so that it has a parent). Read-only.|
|[ParentShowDetail](pivotitem-parentshowdetail-property-excel.md)| **True** if the specified item is showing because one of its parents is showing detail. **False** if the specified item isn't showing because one of its parents is hiding detail. This property is available only if the item is grouped. Read-only **Boolean** .|
|[Position](pivotitem-position-property-excel.md)|Returns or sets a  **Long** value that represents the position of the item in its field, if the item is currently showing.|
|[RecordCount](pivotitem-recordcount-property-excel.md)|Returns the number of records in the PivotTable cache or the number of cache records that contain the specified item. Read-only  **Long** .|
|[ShowDetail](pivotitem-showdetail-property-excel.md)| **True** if the outline is expanded for the specified range (so that the detail of the column or row is visible). The specified range must be a single summary column or row in an outline. Read/write **Variant** . For the **PivotItem** object (or the **Range** object if the range is in a PivotTable report), this property is set to **True** if the item is showing detail.|
|[SourceName](pivotitem-sourcename-property-excel.md)|Returns a  **Variant** value that represents the specified object's name as it appears in the original source data for the specified PivotTable report.|
|[SourceNameStandard](pivotitem-sourcenamestandard-property-excel.md)|Returns a  **String** that represents the PivotTable items' source name in standard English (United States) format settings. Read-only.|
|[StandardFormula](pivotitem-standardformula-property-excel.md)|Returns or sets a  **String** specifying formulas with standard English (United States) formatting. Read/write.|
|[Value](pivotitem-value-property-excel.md)|Returns or sets a  **String** value that represents the name of the specified item in the PivotTable field.|
|[Visible](pivotitem-visible-property-excel.md)|Returns or sets a  **Boolean** value that determines whether the object is visible. Read/write.|

