---
title: UniqueValues Members (Excel)
ms.prod: EXCEL
ms.assetid: 53c161ba-b9ef-e052-2fd3-4c662454c5fc
---


# UniqueValues Members (Excel)
The  **UniqueValues** object uses the **DupeUnique** property to returns or sets an enum that determines whether the rule should look for duplicate or unique values in the range.

The  **UniqueValues** object uses the **DupeUnique** property to returns or sets an enum that determines whether the rule should look for duplicate or unique values in the range.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Delete](uniquevalues-delete-method-excel.md)|Deletes the specified conditional formatting rule object.|
|[ModifyAppliesToRange](uniquevalues-modifyappliestorange-method-excel.md)|Sets the cell range to which this formatting rule applies.|
|[SetFirstPriority](uniquevalues-setfirstpriority-method-excel.md)|Sets the priority value for this conditional formatting rule to "1" so that it will be evaluated before all other rules on the worksheet.|
|[SetLastPriority](uniquevalues-setlastpriority-method-excel.md)|Sets the evaluation order for this conditional formatting rule so it is evaluated after all other rules on the worksheet.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](uniquevalues-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object. Read-only.|
|[AppliesTo](uniquevalues-appliesto-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object specifying the cell range to which the formatting rule is applied.|
|[Borders](uniquevalues-borders-property-excel.md)|Returns a  **[Borders](borders-object-excel.md)** collection that specifies the formatting of cell borders if the conditional formatting rule evaluates to **True** . Read-only.|
|[Creator](uniquevalues-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[DupeUnique](uniquevalues-dupeunique-property-excel.md)|Returns or sets one of the constants of the  **[XlDupeUnique](xldupeunique-enumeration-excel.md)** enumeration, specifying if the conditional format rule is looking for unique or duplicate values.|
|[Font](uniquevalues-font-property-excel.md)|Returns a  **[Font](font-object-excel.md)** object that specifies the font formatting if the conditional formatting rule evaluates to **True** . Read-only.|
|[Interior](uniquevalues-interior-property-excel.md)|Returns an  **[Interior](interior-object-excel.md)** object that specifies a cell's interior attributes for a conditional formatting rule that evaluates to **True** . Read-only.|
|[NumberFormat](uniquevalues-numberformat-property-excel.md)|Returns or sets the number format applied to a cell if the conditional formatting rule evaluates to  **True** . Read/write **Variant** .|
|[Parent](uniquevalues-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[Priority](uniquevalues-priority-property-excel.md)|Returns or sets the priority value of the conditional formatting rule. The priority determines the order of evaluation when multiple conditional formatting rules exist in a worksheet.|
|[PTCondition](uniquevalues-ptcondition-property-excel.md)|Returns a  **Boolean** value indicating if the conditional format is being applied to a PivotTable. Read-only.|
|[ScopeType](uniquevalues-scopetype-property-excel.md)|Returns or sets one of the constants of the  **[XlPivotConditionScope](xlpivotconditionscope-enumeration-excel.md)** enumeration, which determines the scope of the conditional format when it is applied to a PivotTable.|
|[StopIfTrue](uniquevalues-stopiftrue-property-excel.md)|Returns or sets a  **Boolean** value that determines if additional formatting rules on the cell should be evaluated if the current rule evaluates to **True** .|
|[Type](uniquevalues-type-property-excel.md)|Returns one of the constants of the  **[XlFormatConditionType](xlformatconditiontype-enumeration-excel.md)** enumeration, which specifies the type of conditional format. Read-only.|

