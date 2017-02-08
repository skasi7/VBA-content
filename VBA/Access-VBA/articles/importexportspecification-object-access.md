---
title: ImportExportSpecification Object (Access)
keywords: vbaac10.chm13327
f1_keywords:
- vbaac10.chm13327
ms.prod: ACCESS
api_name:
- Access.ImportExportSpecification
ms.assetid: a274faba-6da3-35c5-52fc-3341e8def24a
---


# ImportExportSpecification Object (Access)

Represents a saved import or export operation.


## Remarks

A  **ImportExportSpecification** object contains all the information Access needs to repeat an import or export operation without your having to provide any input. For example, an import specification that imports data from a Microsoft Office Excel 2007 workbook stores the name of the source Excel file, the name of the destination database, and other details, such as whether you appended to or created a new table, primary key information, field names, and so on.

Use the  **[Add](http://msdn.microsoft.com/library/importexportspecifications-add-method-access%28Office.15%29.aspx)** method of the **[ImportExportSpecifications](http://msdn.microsoft.com/library/importexportspecifications-object-access%28Office.15%29.aspx)** collection to create a new **ImportExportSpecification** object.

Use the  **[Execute](http://msdn.microsoft.com/library/importexportspecification-execute-method-access%28Office.15%29.aspx)** method to run saved import or export operation.


## Methods



|**Name**|
|:-----|
|[Delete](http://msdn.microsoft.com/library/importexportspecification-delete-method-access%28Office.15%29.aspx)|
|[Execute](http://msdn.microsoft.com/library/importexportspecification-execute-method-access%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/importexportspecification-application-property-access%28Office.15%29.aspx)|
|[Description](http://msdn.microsoft.com/library/importexportspecification-description-property-access%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/importexportspecification-name-property-access%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/importexportspecification-parent-property-access%28Office.15%29.aspx)|
|[XML](http://msdn.microsoft.com/library/importexportspecification-xml-property-access%28Office.15%29.aspx)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/object-model-access-vba-reference%28Office.15%29.aspx)
[ImportExportSpecification Object Members](http://msdn.microsoft.com/library/importexportspecification-members-access%28Office.15%29.aspx)
