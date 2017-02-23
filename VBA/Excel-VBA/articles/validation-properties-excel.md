---
title: Validation Properties (Excel)
ms.prod: EXCEL
ms.assetid: 8af57426-9cf0-4209-bac5-a6166d20ccf3
---


# Validation Properties (Excel)

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[AlertStyle](validation-alertstyle-property-excel.md)|Returns the validation alert style. Read-only  **[XlDVAlertStyle](xldvalertstyle-enumeration-excel.md)** .|
|[Application](validation-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[Creator](validation-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[ErrorMessage](validation-errormessage-property-excel.md)|Returns or sets the data validation error message. Read/write  **String** .|
|[ErrorTitle](validation-errortitle-property-excel.md)|Returns or sets the title of the data-validation error dialog box. Read/write  **String** .|
|[Formula1](validation-formula1-property-excel.md)|Returns the value or expression associated with the conditional format or data validation. Can be a constant value, a string value, a cell reference, or a formula. Read-only  **String** .|
|[Formula2](validation-formula2-property-excel.md)|Returns the value or expression associated with the second part of a conditional format or data validation. Used only when the data validation conditional format  **[Operator](validation-operator-property-excel.md)** property is **xlBetween** or **xlNotBetween** . Can be a constant value, a string value, a cell reference, or a formula. Read-only **String** .|
|[IgnoreBlank](validation-ignoreblank-property-excel.md)| **True** if blank values are permitted by the range data validation. Read/write **Boolean** .|
|[IMEMode](validation-imemode-property-excel.md)|Returns or sets the description of the Japanese input rules. Can be one of the  **[XlIMEMode](xlimemode-enumeration-excel.md)** constants listed in the following table. Read/write **Long** .|
|[InCellDropdown](validation-incelldropdown-property-excel.md)| **True** if data validation displays a drop-down list that contains acceptable values. Read/write **Boolean** .|
|[InputMessage](validation-inputmessage-property-excel.md)|Returns or sets the data validation input message. Read/write  **String** .|
|[InputTitle](validation-inputtitle-property-excel.md)|Returns or sets the title of the data-validation input dialog box. Read/write  **String** .|
|[Operator](validation-operator-property-excel.md)|Returns a  **Long** value that represents the operator for the data validation.|
|[Parent](validation-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[ShowError](validation-showerror-property-excel.md)| **True** if the data validation error message will be displayed whenever the user enters invalid data. Read/write **Boolean** .|
|[ShowInput](validation-showinput-property-excel.md)| **True** if the data validation input message will be displayed whenever the user selects a cell in the data validation range. Read/write **Boolean** .|
|[Type](validation-type-property-excel.md)|Returns a  **Long** value, containing a **[XlDVType](xldvtype-enumeration-excel.md)** constant, that represents the data type validation for a range.|
|[Value](validation-value-property-excel.md)|Returns a  **Boolean** value that indicates if all the validation criteria are met (that is, if the range contains valid data).|

