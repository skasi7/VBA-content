---
title: IconSetCondition Members (Excel)
ms.prod: EXCEL
ms.assetid: 5ea20648-be46-7b8b-be31-368fc98329ab
---


# IconSetCondition Members (Excel)
Represents an icon set conditional formatting rule.

Represents an icon set conditional formatting rule.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Delete](iconsetcondition-delete-method-excel.md)|Deletes the specified conditional formatting rule object.|
|[ModifyAppliesToRange](iconsetcondition-modifyappliestorange-method-excel.md)|Sets the cell range to which this formatting rule applies.|
|[SetFirstPriority](iconsetcondition-setfirstpriority-method-excel.md)|Sets the priority value for this conditional formatting rule to "1" so that it will be evaluated before all other rules on the worksheet.|
|[SetLastPriority](iconsetcondition-setlastpriority-method-excel.md)|Sets the evaluation order for this conditional formatting rule so it is evaluated after all other rules on the worksheet.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](iconsetcondition-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object. Read-only.|
|[AppliesTo](iconsetcondition-appliesto-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object specifying the cell range to which the formatting rule is applied.|
|[Creator](iconsetcondition-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[Formula](iconsetcondition-formula-property-excel.md)|Returns or sets a  **String** representing a formula, which determines the values to which the icon set will be applied.|
|[IconCriteria](iconsetcondition-iconcriteria-property-excel.md)|Returns an  **[IconCriteria](iconcriteria-object-excel.md)** collection, which represents the set of criteria for an icon set conditional formatting rule.|
|[IconSet](iconsetcondition-iconset-property-excel.md)|Returns or sets an  **[IconSets](iconsets-object-excel.md)** collection, which specifies the icon set used in the conditional format.|
|[Parent](iconsetcondition-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[PercentileValues](iconsetcondition-percentilevalues-property-excel.md)|Returns or sets a  **Boolean** value indicating if the thresholds for an icon set conditional format are determined by using percentiles.|
|[Priority](iconsetcondition-priority-property-excel.md)|Returns or sets the priority value of the conditional formatting rule. The priority determines the order of evaluation when multiple conditional formatting rules exist in a worksheet.|
|[PTCondition](iconsetcondition-ptcondition-property-excel.md)|Returns a  **Boolean** value indicating if the conditional format is being applied to a PivotTable. Read-only.|
|[ReverseOrder](iconsetcondition-reverseorder-property-excel.md)|Returns or sets a  **Boolean** value indicating if the order of icons is reversed for an icon set.|
|[ScopeType](iconsetcondition-scopetype-property-excel.md)|Returns or sets one of the constants of the  **[XlPivotConditionScope](xlpivotconditionscope-enumeration-excel.md)** enumeration, which determines the scope of the conditional format when it is applied to a PivotTable.|
|[ShowIconOnly](iconsetcondition-showicononly-property-excel.md)|Returns or sets a  **Boolean** value indicating if only the icon is displayed for an icon set conditional format.|
|[StopIfTrue](iconsetcondition-stopiftrue-property-excel.md)|Returns or sets a  **Boolean** value that determines if additional formatting rules on the cell should be evaluated if the current rule evaluates to **True** .|
|[Type](iconsetcondition-type-property-excel.md)|Returns one of the constants of the  **[XlFormatConditionType](xlformatconditiontype-enumeration-excel.md)** enumeration, which specifies the type of conditional format. Read-only.|

