---
title: Rows Object (PowerPoint)
keywords: vbapp10.chm625000
f1_keywords:
- vbapp10.chm625000
ms.prod: POWERPOINT
ms.assetid: 9a72b6bb-2aec-e37b-f1a2-005f910e1335
---


# Rows Object (PowerPoint)

A collection of  **[Row](http://msdn.microsoft.com/library/row-object-powerpoint%28Office.15%29.aspx)** objects that represent the rows in a table.


## Example

Use the [Rows](http://msdn.microsoft.com/library/table-rows-property-powerpoint%28Office.15%29.aspx)property to return the  **Rows** collection. This example changes the height of all rows in the specified table to 160 points.


```
Dim i As Integer

With ActivePresentation.Slides(2).Shapes(4).Table

    For i = 1 To .Rows.Count

        .Rows.Height = 160

    Next i

End With
```

Use the [Add](http://msdn.microsoft.com/library/rows-add-method-powerpoint%28Office.15%29.aspx)method to add a row to a table. This example inserts a row before the second row in the referenced table.




```
ActivePresentation.Slides(2).Shapes(5).Table.Rows.Add (2)
```

Use  **Rows** (index), where index is a number that represents the position of the row in the table, to return a single **Row** object. This example deletes the first row from the table in shape five on slide two.




```
ActivePresentation.Slides(2).Shapes(5).Table.Rows(1).Delete
```


## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/rows-add-method-powerpoint%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/rows-item-method-powerpoint%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/rows-application-property-powerpoint%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/rows-count-property-powerpoint%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/rows-parent-property-powerpoint%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/object-model-powerpoint-vba-reference%28Office.15%29.aspx)
