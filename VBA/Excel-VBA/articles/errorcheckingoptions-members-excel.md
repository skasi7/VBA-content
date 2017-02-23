---
title: ErrorCheckingOptions Members (Excel)
ms.prod: EXCEL
ms.assetid: 257ede5e-bbc2-2da7-d2e1-f62ff0f02512
---


# ErrorCheckingOptions Members (Excel)
Represents the error-checking options for an application.

Represents the error-checking options for an application.


## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](errorcheckingoptions-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[BackgroundChecking](errorcheckingoptions-backgroundchecking-property-excel.md)|Alerts the user for all cells that violate enabled error-checking rules. When this property is set to  **True** (default), the **AutoCorrect Options** button appears next to all cells that violate enabled errors. **False** disables background checking for errors. Read/write **Boolean** .|
|[Creator](errorcheckingoptions-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[EmptyCellReferences](errorcheckingoptions-emptycellreferences-property-excel.md)|When set to  **True** (default), Microsoft Excel identifies, with an **AutoCorrect Options** button, selected cells containing formulas that refer to empty cells. **False** disables empty cell reference checking. Read/write **Boolean** .|
|[EvaluateToError](errorcheckingoptions-evaluatetoerror-property-excel.md)|When set to  **True** (default), Microsoft Excel identifies, with an **AutoCorrect Options** button, selected cells that contain formulas evaluating to an error. **False** disables error checking for cells that evaluate to an error value. Read/write **Boolean** .|
|[InconsistentFormula](errorcheckingoptions-inconsistentformula-property-excel.md)|When set to  **True** (default), Microsoft Excel identifies cells containing an inconsistent formula in a region. **False** disables the inconsistent formula check. Read/write **Boolean** .|
|[InconsistentTableFormula](errorcheckingoptions-inconsistenttableformula-property-excel.md)|Returns  **True** if the table formula is inconsistent. Read/write **Boolean** .|
|[IndicatorColorIndex](errorcheckingoptions-indicatorcolorindex-property-excel.md)|Returns or sets the color of the indicator for error checking options. Read/write  **[XlColorIndex](xlcolorindex-enumeration-excel.md)** .|
|[ListDataValidation](errorcheckingoptions-listdatavalidation-property-excel.md)|A  **Boolean** value that is **True** if data validation is enabled in a list. Read/write **Boolean** .|
|[NumberAsText](errorcheckingoptions-numberastext-property-excel.md)|When set to  **True** (default), Microsoft Excel identifies, with an **AutoCorrect Options** button, selected cells that contain numbers written as text. **False** disables error checking for numbers written as text. Read/write **Boolean** .|
|[OmittedCells](errorcheckingoptions-omittedcells-property-excel.md)|When set to  **True** (default), Microsoft Excel identifies, with an AutoCorrect Options button, the selected cells that contain formulas referring to a range that omits adjacent cells that could be included. **False** disables error checking for omitted cells. Read/write **Boolean** .|
|[Parent](errorcheckingoptions-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[TextDate](errorcheckingoptions-textdate-property-excel.md)|When set to  **True** (default), Microsoft Excel identifies, with an **AutoCorrect Options** button, cells that contain a text date with a two-digit year. **False** disables error checking for cells containing a text date with a two-digit year. Read/write **Boolean** .|
|[UnlockedFormulaCells](errorcheckingoptions-unlockedformulacells-property-excel.md)|When set to  **True** (default), Microsoft Excel identifies selected cells that are unlocked and contain a formula. **False** disables error checking for unlocked cells that contain formulas. Read/write **Boolean** .|

