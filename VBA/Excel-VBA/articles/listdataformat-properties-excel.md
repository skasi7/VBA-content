---
title: ListDataFormat Properties (Excel)
ms.prod: EXCEL
ms.assetid: 7a041aab-e889-4519-baa7-340104c4a457
---


# ListDataFormat Properties (Excel)

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[AllowFillIn](listdataformat-allowfillin-property-excel.md)| Returns a **Boolean** value indicating whether users can provide their own data for cells in a column (rather than being restricted to a list of values) for those columns that supply a list of values. Returns **False** for lists that are not linked to a SharePoint site. Also returns **False** if the column is not a specified as choice or multi-choice. Read-only **Boolean** .|
|[Application](listdataformat-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[Choices](listdataformat-choices-property-excel.md)| Returns an **Array** of **String** values that contains the choices offered to the user by the **ListLookUp** , **ChoiceMulti** , and **Choice** data types of the **[DefaultValue](listdataformat-defaultvalue-property-excel.md)** property. Read-only **Variant** .|
|[Creator](listdataformat-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[DecimalPlaces](listdataformat-decimalplaces-property-excel.md)|Returns a  **Long** value that represents the number of decimal places to show for the numbers in the **[ListColumn](listcolumn-object-excel.md)** object. Read-only **Long** .|
|[DefaultValue](listdataformat-defaultvalue-property-excel.md)| Returns **Variant** representing the default data type value for a new row in a column. The **Nothing** object is returned when the schema does not specify a default value. Read-only **Variant** .|
|[IsPercent](listdataformat-ispercent-property-excel.md)|Returns a  **Boolean** value. Returns **True** only if the number data for the **[ListColumn](listcolumn-object-excel.md)** object will be shown in percentage formatting. Read-only **Boolean** . Read-only.|
|[lcid](listdataformat-lcid-property-excel.md)|Returns a  **Long** value that represents the LCID for the **[ListColumn](listcolumn-object-excel.md)** object that is specified in the schema definition. Read-only **Long** .|
|[MaxCharacters](listdataformat-maxcharacters-property-excel.md)|Returns a  **Long** containing the maximum number of characters allowed in the **[ListColumn](listcolumn-object-excel.md)** object if the **[Type](listdataformat-type-property-excel.md)** property is set to **xlListDataTypeText** or **xlListDataTypeMultiLineText** . Read-only **Long** .|
|[MaxNumber](listdataformat-maxnumber-property-excel.md)|Returns a  **Variant** containing the maximum value allowed in this field in the list column. Read-only **Variant** .|
|[MinNumber](listdataformat-minnumber-property-excel.md)| Returns a **Variant** containing the minimum value allowed in this field in the list column. This can be a negative floating point number. Read-only **Variant** .|
|[Parent](listdataformat-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[ReadOnly](listdataformat-readonly-property-excel.md)| Returns **True** if the object has been opened as read-only. Read-only **Boolean** .|
|[Required](listdataformat-required-property-excel.md)| Returns a **Boolean** value indicating whether the schema definition of a column requires data before the row is committed. Read-only **Boolean** .|
|[Type](listdataformat-type-property-excel.md)|Returns an  **[XlListDataType](xllistdatatype-enumeration-excel.md)** value that represents the data type of the list column.|

