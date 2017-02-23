---
title: Databar Members (Excel)
ms.prod: EXCEL
ms.assetid: 137f7e88-bb61-48a3-d2cb-76a8282cd62e
---


# Databar Members (Excel)
Represents a data bar conditional formating rule. Applying a data bar to a range helps you see the value of a cell relative to other cells.

Represents a data bar conditional formating rule. Applying a data bar to a range helps you see the value of a cell relative to other cells.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Delete](databar-delete-method-excel.md)|Deletes the specified conditional formatting rule object.|
|[ModifyAppliesToRange](databar-modifyappliestorange-method-excel.md)|Sets the cell range to which this formatting rule applies.|
|[SetFirstPriority](databar-setfirstpriority-method-excel.md)|Sets the priority value for this conditional formatting rule to "1" so that it will be evaluated before all other rules on the worksheet.|
|[SetLastPriority](databar-setlastpriority-method-excel.md)|Sets the evaluation order for this conditional formatting rule so it is evaluated after all other rules on the worksheet.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](databar-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object. Read-only.|
|[AppliesTo](databar-appliesto-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object specifying the cell range to which the formatting rule is applied.|
|[AxisColor](databar-axiscolor-property-excel.md)|Returns the color of the axis for cells with conditional formatting as data bars. Read-only|
|[AxisPosition](databar-axisposition-property-excel.md)|Returns or sets the position of the axis of the data bars specified by a conditional formatting rule. Read/write|
|[BarBorder](databar-barborder-property-excel.md)|Returns an object that specifies the border of a data bar. Read-only|
|[BarColor](databar-barcolor-property-excel.md)|Returns a  **[FormatColor](formatcolor-object-excel.md)** object that you can use to modify the color of the bars in a data bar conditional format.|
|[BarFillType](databar-barfilltype-property-excel.md)|Returns or sets how a data bar is filled with color. Read/write|
|[Creator](databar-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[Direction](databar-direction-property-excel.md)|Returns or sets the direction the databar is displayed. Read/write|
|[Formula](databar-formula-property-excel.md)|Returns or sets a  **String** representing a formula, which determines the values to which the data bar will be applied.|
|[MaxPoint](databar-maxpoint-property-excel.md)|Returns a  **[ConditionValue](conditionvalue-object-excel.md)** object that specifies how the longest bar is evaluated for a data bar conditional format.|
|[MinPoint](databar-minpoint-property-excel.md)|Returns a  **[ConditionValue](conditionvalue-object-excel.md)** object that specifies how the shortest bar is evaluated for a data bar conditional format.|
|[NegativeBarFormat](databar-negativebarformat-property-excel.md)|Returns the  **[NegativeBarFormat](negativebarformat-object-excel.md)** object associated with a data bar conditional formatting rule. Read-only|
|[Parent](databar-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[PercentMax](databar-percentmax-property-excel.md)|Returns or sets a  **Long** value that specifies the length of the longest data bar as a percentage of cell width.|
|[PercentMin](databar-percentmin-property-excel.md)|Returns or sets a  **Long** value that specifies the length of the shortest data bar as a percentage of cell width.|
|[Priority](databar-priority-property-excel.md)|Returns or sets the priority value of the conditional formatting rule. The priority determines the order of evaluation when multiple conditional formatting rules exist in a worksheet.|
|[PTCondition](databar-ptcondition-property-excel.md)|Returns a  **Boolean** value indicating if the conditional format is being applied to a PivotTable. Read-only.|
|[ScopeType](databar-scopetype-property-excel.md)|Returns or sets one of the constants of the  **[XlPivotConditionScope](xlpivotconditionscope-enumeration-excel.md)** enumeration, which determines the scope of the conditional format when it is applied to a PivotTable.|
|[ShowValue](databar-showvalue-property-excel.md)|Returns or sets a  **Boolean** value that specifies if the value in the cell is displayed if the data bar conditional format is applied to the range.|
|[StopIfTrue](databar-stopiftrue-property-excel.md)|Returns or sets a  **Boolean** value that determines if additional formatting rules on the cell should be evaluated if the current rule evaluates to **True** .|
|[Type](databar-type-property-excel.md)|Returns one of the constants of the  **[XlFormatConditionType](xlformatconditiontype-enumeration-excel.md)** enumeration, which specifies the type of conditional format. Read-only.|

