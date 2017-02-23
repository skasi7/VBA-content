---
title: FormatCondition Members (Excel)
ms.prod: EXCEL
ms.assetid: 8f4bebce-0bf4-03de-62f0-4454ea699c5f
---


# FormatCondition Members (Excel)
Represents a conditional format.

Represents a conditional format.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Delete](formatcondition-delete-method-excel.md)|Deletes the object.|
|[Modify](formatcondition-modify-method-excel.md)|Modifies an existing conditional format.|
|[ModifyAppliesToRange](formatcondition-modifyappliestorange-method-excel.md)|Sets the cell range to which this formatting rule applies.|
|[SetFirstPriority](formatcondition-setfirstpriority-method-excel.md)|Sets the priority value for this conditional formatting rule to "1" so that it will be evaluated before all other rules on the worksheet.|
|[SetLastPriority](formatcondition-setlastpriority-method-excel.md)|Sets the evaluation order for this conditional formatting rule so it is evaluated after all other rules on the worksheet.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](formatcondition-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[AppliesTo](formatcondition-appliesto-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object specifying the cell range to which the formatting rule is applied.|
|[Borders](formatcondition-borders-property-excel.md)|Returns a  **[Borders](borders-object-excel.md)** collection that represents the borders of a style or a range of cells (including a range defined as part of a conditional format).|
|[Creator](formatcondition-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[DateOperator](formatcondition-dateoperator-property-excel.md)|Specifies the Date operator used in the format condition. Read/write.|
|[Font](formatcondition-font-property-excel.md)|Returns a  **[Font](font-object-excel.md)** object that represents the font of the specified object.|
|[Formula1](formatcondition-formula1-property-excel.md)|Returns the value or expression associated with the conditional format or data validation. Can be a constant value, a string value, a cell reference, or a formula. Read-only  **String** .|
|[Formula2](formatcondition-formula2-property-excel.md)|Returns the value or expression associated with the second part of a conditional format or data validation. Used only when the data validation conditional format  **[Operator](formatcondition-operator-property-excel.md)** property is **xlBetween** or **xlNotBetween** . Can be a constant value, a string value, a cell reference, or a formula. Read-only **String** .|
|[Interior](formatcondition-interior-property-excel.md)|Returns an  **[Interior](interior-object-excel.md)** object that represents the interior of the specified object.|
|[NumberFormat](formatcondition-numberformat-property-excel.md)|Returns or sets the number format applied to a cell if the conditional formatting rule evaluates to  **True** . Read/write **Variant** .|
|[Operator](formatcondition-operator-property-excel.md)|Returns a  **Long** value that represents the operator for the conditional format.|
|[Parent](formatcondition-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[Priority](formatcondition-priority-property-excel.md)|Returns or sets the priority value of the conditional formatting rule. The priority determines the order of evaluation when multiple conditional formatting rules exist in a worksheet.|
|[PTCondition](formatcondition-ptcondition-property-excel.md)|Returns a  **Boolean** value indicating if the conditional format is being applied to a PivotTable. Read-only.|
|[ScopeType](formatcondition-scopetype-property-excel.md)|Returns or sets one of the constants of the  **[XlPivotConditionScope](xlpivotconditionscope-enumeration-excel.md)** enumeration, which determines the scope of the conditional format when it is applied to a PivotTable.|
|[StopIfTrue](formatcondition-stopiftrue-property-excel.md)|Returns or sets a  **Boolean** value that determines if additional formatting rules on the cell should be evaluated if the current rule evaluates to **True** .|
|[Text](formatcondition-text-property-excel.md)|Returns or sets a  **String** value specifying the text string used by the conditional formatting rule.|
|[TextOperator](formatcondition-textoperator-property-excel.md)|Returns or sets one of the constants of the  **[XlContainsOperator](xlcontainsoperator-enumeration-excel.md)** enumeration, specifying the text search performed by the conditional formatting rule.|
|[Type](formatcondition-type-property-excel.md)|Returns a  **Long** value, containing a **[xlFormatConditionType](xlformatconditiontype-enumeration-excel.md)** value, that represents the object type.|

