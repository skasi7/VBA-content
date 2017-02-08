---
title: AboveAverage Members (Excel)
ms.prod: EXCEL
ms.assetid: 85828a41-ce2a-4979-8918-3adaed2f5661
---


# AboveAverage Members (Excel)
Represents an above average visual of a conditional formatting rule. Applying a color or fill to a range or selection to help you see the value of a cells relative to other cells.

Represents an above average visual of a conditional formatting rule. Applying a color or fill to a range or selection to help you see the value of a cells relative to other cells.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Delete](aboveaverage-delete-method-excel.md)|Deletes the specified conditional formatting rule object.|
|[ModifyAppliesToRange](aboveaverage-modifyappliestorange-method-excel.md)|Sets the cell range to which this formatting rule applies.|
|[SetFirstPriority](aboveaverage-setfirstpriority-method-excel.md)|Sets the priority value for this conditional formatting rule to "1" so that it will be evaluated before all other rules on the worksheet.|
|[SetLastPriority](aboveaverage-setlastpriority-method-excel.md)|Sets the evaluation order for this conditional formatting rule so it is evaluated after all other rules on the worksheet.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[AboveBelow](aboveaverage-abovebelow-property-excel.md)|Returns or sets one of the constants of the  **[XlAboveBelow](xlabovebelow-enumeration-excel.md)** enumeration, specifying if the conditional formatting rule looks for cell values above or below the range average or standard deviation.|
|[Application](aboveaverage-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object. Read-only.|
|[AppliesTo](aboveaverage-appliesto-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object specifying the cell range to which the formatting rule is applied.|
|[Borders](aboveaverage-borders-property-excel.md)|Returns a  **[Borders](borders-object-excel.md)** collection that specifies the formatting of cell borders if the conditional formatting rule evaluates to **True** . Read-only.|
|[CalcFor](aboveaverage-calcfor-property-excel.md)|Returns or sets one of the constants of the  **[XlCalcFor](xlcalcfor-enumeration-excel.md)** enumeration, which specifies the scope of data to be evaluated for the conditional format in a PivotTable report.|
|[Creator](aboveaverage-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[Font](aboveaverage-font-property-excel.md)|Returns a  **[Font](font-object-excel.md)** object that specifies the font formatting if the conditional formatting rule evaluates to **True** . Read-only.|
|[Interior](aboveaverage-interior-property-excel.md)|Returns an  **[Interior](interior-object-excel.md)** object that specifies a cell's interior attributes for a conditional formatting rule that evaluates to **True** . Read-only.|
|[NumberFormat](aboveaverage-numberformat-property-excel.md)|Returns or sets the number format applied to a cell if the conditional formatting rule evaluates to  **True** . Read/write **Variant** .|
|[NumStdDev](aboveaverage-numstddev-property-excel.md)|Returns or sets the numeric standard deviation for an  **AboveAverage** object. Read/write **Long** .|
|[Parent](aboveaverage-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[Priority](aboveaverage-priority-property-excel.md)|Returns or sets the priority value of the conditional formatting rule. The priority determines the order of evaluation when multiple conditional formatting rules exist in a worksheet.|
|[PTCondition](aboveaverage-ptcondition-property-excel.md)|Returns a  **Boolean** value indicating if the conditional format is being applied to a PivotTable. Read-only.|
|[ScopeType](aboveaverage-scopetype-property-excel.md)|Returns or sets one of the constants of the  **[XlPivotConditionScope](xlpivotconditionscope-enumeration-excel.md)** enumeration, which determines the scope of the conditional format when it is applied to a PivotTable.|
|[StopIfTrue](aboveaverage-stopiftrue-property-excel.md)|Returns or sets a  **Boolean** value that determines if additional formatting rules on the cell should be evaluated if the current rule evaluates to **True** .|
|[Type](aboveaverage-type-property-excel.md)|Returns one of the constants of the  **[XlFormatConditionType](xlformatconditiontype-enumeration-excel.md)** enumeration, which specifies the type of conditional format. Read-only.|

