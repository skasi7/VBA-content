---
title: RecentFiles Object (Excel)
keywords: vbaxl10.chm171072
f1_keywords:
- vbaxl10.chm171072
ms.prod: EXCEL
api_name:
- Excel.RecentFiles
ms.assetid: e33ae942-0444-0631-be08-386366b6ebdb
---


# RecentFiles Object (Excel)

Represents the list of recently used files.


## Remarks

 Each file is represented by a **[RecentFile](recentfile-object-excel.md)** object.


## Example

Use the  **[RecentFiles](application-recentfiles-property-excel.md)** property to return the **RecentFiles** collection. The following example sets the maximum number of files in the list of recently used files.


```vb
Application.RecentFiles.Maximum = 6
```


## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/object-model-excel-vba-reference%28Office.15%29.aspx)


