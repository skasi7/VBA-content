---
title: Cell Object (PowerPoint)
keywords: vbapp10.chm628000
f1_keywords:
- vbapp10.chm628000
ms.prod: POWERPOINT
ms.assetid: e89e5d69-33b1-d7b1-0a6c-4dfd8b676977
---


# Cell Object (PowerPoint)

Represents a table cell. The  **Cell** object is a member of the **[CellRange](http://msdn.microsoft.com/library/cellrange-object-powerpoint%28Office.15%29.aspx)** collection. The **CellRange** collection represents all the cells in the specified column or row. To use the **CellRange** collection, use the **Cells** keyword.


## Remarks

You cannot programmatically add cells to or delete cells from a PowerPoint table. Use the  **Add** method of the **Columns** or **Rows** collections to add a column or row to a table. Use the **Delete** method of the **Columns** or **Rows** collections to delete a column or row from a table.


## Example

Use  **Cell** (row, column), where row is the row number and column is the column number, or **Cells** (index), where index is the number of the cell in the specified row or column, to return a single **Cell** object. Cells are numbered from left to right in rows and from top to bottom in columns. With right-to-left language settings, this scheme is reversed. The following example merges the first two cells in row one of the table in shape five on slide two.


```
With ActivePresentation.Slides(2).Shapes(5).Table

    .Cell(1, 1).Merge MergeTo:=.Cell(1, 2)

End With
```

This example sets the bottom border for cell one in the first column of the table to a dashed line style.




```
With ActivePresentation.Slides(2).Shapes(5).Table.Columns(1) _

        .Cells(1)

    .Borders(ppBorderBottom).DashStyle = msoLineDash

End With
```

Use the [Shape](http://msdn.microsoft.com/library/cell-shape-property-powerpoint%28Office.15%29.aspx)property to access the  **Shape** object and to manipulate the contents of each cell. This example deletes the text in the first cell (row 1, column 1), inserts new text, and then sets the width of the entire column to 110 points.




```
With ActivePresentation.Slides(2).Shapes(5).Table.Cell(1, 1)

    .Shape.TextFrame.TextRange.Delete

    .Shape.TextFrame.TextRange.Text = "Rooster"

    .Parent.Columns(1).Width = 110

End With
```


## Methods



|**Name**|
|:-----|
|[Merge](http://msdn.microsoft.com/library/cell-merge-method-powerpoint%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/cell-select-method-powerpoint%28Office.15%29.aspx)|
|[Split](http://msdn.microsoft.com/library/cell-split-method-powerpoint%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/cell-application-property-powerpoint%28Office.15%29.aspx)|
|[Borders](http://msdn.microsoft.com/library/cell-borders-property-powerpoint%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/cell-parent-property-powerpoint%28Office.15%29.aspx)|
|[Selected](http://msdn.microsoft.com/library/cell-selected-property-powerpoint%28Office.15%29.aspx)|
|[Shape](http://msdn.microsoft.com/library/cell-shape-property-powerpoint%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/object-model-powerpoint-vba-reference%28Office.15%29.aspx)
