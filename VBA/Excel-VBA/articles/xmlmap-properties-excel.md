---
title: XmlMap Properties (Excel)
ms.prod: EXCEL
ms.assetid: 421f12ed-b650-401f-9b12-99ce30c6d183
---


# XmlMap Properties (Excel)

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[AdjustColumnWidth](xmlmap-adjustcolumnwidth-property-excel.md)| **True** if the column widths are automatically adjusted for the best fit each time you refresh the specified XML map. **False** if the column widths aren't automatically adjusted with each refresh. The default value is **True** . Read/write **Boolean** .|
|[AppendOnImport](xmlmap-appendonimport-property-excel.md)| **True** if you want to append new rows to XML lists that are bound to the specified schema map when you are importing new data or refreshing an existing connection. **False** if you want to overwrite the contents of cells that are bound to the specified schema map when you are importing new data or refreshing an existing connection. The default value is **False** . Read/write **Boolean** .|
|[Application](xmlmap-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[Creator](xmlmap-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[DataBinding](xmlmap-databinding-property-excel.md)|Returns an  **[XmlDataBinding](xmldatabinding-object-excel.md)** object that represents the binding associated with the specified schema map. Read-only.|
|[IsExportable](xmlmap-isexportable-property-excel.md)|Returns  **True** if Microsoft Excel can use the **[XPath](xpath-object-excel.md)** objects in the specified schema map to export XML data and all XML lists mapped to the specified schema map can be exported.|
|[Name](xmlmap-name-property-excel.md)|Returns or sets a  **String** value that represents the friendly name used to uniquely identify a mapping in the workbook.|
|[Parent](xmlmap-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[PreserveColumnFilter](xmlmap-preservecolumnfilter-property-excel.md)|Returns or sets whether filtering is preserved when the specified XML map is refreshed. Read/write  **Boolean** .|
|[PreserveNumberFormatting](xmlmap-preservenumberformatting-property-excel.md)| **True** if number formatting on cells mapped to the specified XML schema map will be preserved when the schema map is refreshed. The default value is **False** . Read/write **Boolean** .|
|[RootElementName](xmlmap-rootelementname-property-excel.md)| Returns a **String** that represents the name of the root element for the specified XML schema map. Read-only.|
|[RootElementNamespace](xmlmap-rootelementnamespace-property-excel.md)|Returns an  **[XmlNamespace](xmlnamespace-object-excel.md)** object that represents the root element for the specified XML schema map. Read-only.|
|[SaveDataSourceDefinition](xmlmap-savedatasourcedefinition-property-excel.md)| **True** if the data source definition of the specified XML schema map is saved with the workbook. The default value is **True** . Read/write **Boolean** .|
|[Schemas](xmlmap-schemas-property-excel.md)| Returns an **[XmlSchemas](xmlschemas-object-excel.md)** collection that represents the schemas that the specified **[XmlMap](xmlmap-object-excel.md)** object contains. Read-only.|
|[ShowImportExportValidationErrors](xmlmap-showimportexportvalidationerrors-property-excel.md)| Returns or sets whether to display a dialog box that details schema-validation errors when data is imported or exported through the specified XML schema map. The default value is **False** . Read/write **Boolean** .|
|[WorkbookConnection](xmlmap-workbookconnection-property-excel.md)|Retuns a new connection for the specified  **XMLMap** object. Read-only.|

