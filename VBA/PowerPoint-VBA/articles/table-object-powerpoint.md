---
title: Table Object (PowerPoint)
keywords: vbapp10.chm622000
f1_keywords:
- vbapp10.chm622000
ms.prod: POWERPOINT
ms.assetid: ebbbca9f-4591-10ce-3c74-33b46a3b7cdf
---


# Table Object (PowerPoint)

Represents a table shape on a slide. The  **Table** object is a member of the **Shapes** collection. The **Table** object contains the **[Columns](http://msdn.microsoft.com/library/columns-object-powerpoint%28Office.15%29.aspx)** collection and the **[Rows](rows-object-powerpoint.md)** collection.


## Example

Use  **Shapes** (index), where index is a number, to return a shape containing a table. Use the[HasTable](http://msdn.microsoft.com/library/shape-hastable-property-powerpoint%28Office.15%29.aspx)property to see if a shape contains a table. This example walks through the shapes on slide one, checks to see if each shape has a table, and then sets the mouse click action for each table shape to advance to the next slide.


```
With ActivePresentation.Slides(2).Shapes

    For i = 1 To .Count

        If .Item(i).HasTable Then

            .Item(i).ActionSettings(ppMouseClick) _

                .Action = ppActionNextSlide

        End If

    Next

End With
```

Use the [Cell](http://msdn.microsoft.com/library/table-cell-method-powerpoint%28Office.15%29.aspx)method of the  **Table** object to access the contents of each cell. This example inserts the text "Cell 1" in the first cell of the table in shape five on slide three.




```
ActivePresentation.Slides(3).Shapes(5).Table _

    .Cell(1, 1).Shape.TextFrame.TextRange _

    .Text = "Cell 1"
```

Use the [AddTable](http://msdn.microsoft.com/library/shapes-addtable-method-powerpoint%28Office.15%29.aspx)method to add a table to a slide. This example adds a 3x3 table on slide two in the active presentation.




```
ActivePresentation.Slides(2).Shapes.AddTable(3, 3)
```


## Methods



|**Name**|
|:-----|
|[ApplyStyle](http://msdn.microsoft.com/library/table-applystyle-method-powerpoint%28Office.15%29.aspx)|
|[Cell](http://msdn.microsoft.com/library/table-cell-method-powerpoint%28Office.15%29.aspx)|
|[ScaleProportionally](http://msdn.microsoft.com/library/table-scaleproportionally-method-powerpoint%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AlternativeText](http://msdn.microsoft.com/library/table-alternativetext-property-powerpoint%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/table-application-property-powerpoint%28Office.15%29.aspx)|
|[Background](http://msdn.microsoft.com/library/table-background-property-powerpoint%28Office.15%29.aspx)|
|[Columns](http://msdn.microsoft.com/library/table-columns-property-powerpoint%28Office.15%29.aspx)|
|[FirstCol](http://msdn.microsoft.com/library/table-firstcol-property-powerpoint%28Office.15%29.aspx)|
|[FirstRow](http://msdn.microsoft.com/library/table-firstrow-property-powerpoint%28Office.15%29.aspx)|
|[HorizBanding](http://msdn.microsoft.com/library/table-horizbanding-property-powerpoint%28Office.15%29.aspx)|
|[LastCol](http://msdn.microsoft.com/library/table-lastcol-property-powerpoint%28Office.15%29.aspx)|
|[LastRow](http://msdn.microsoft.com/library/table-lastrow-property-powerpoint%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/table-parent-property-powerpoint%28Office.15%29.aspx)|
|[Rows](http://msdn.microsoft.com/library/table-rows-property-powerpoint%28Office.15%29.aspx)|
|[Style](http://msdn.microsoft.com/library/table-style-property-powerpoint%28Office.15%29.aspx)|
|[TableDirection](http://msdn.microsoft.com/library/table-tabledirection-property-powerpoint%28Office.15%29.aspx)|
|[Title](http://msdn.microsoft.com/library/table-title-property-powerpoint%28Office.15%29.aspx)|
|[VertBanding](http://msdn.microsoft.com/library/table-vertbanding-property-powerpoint%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/object-model-powerpoint-vba-reference%28Office.15%29.aspx)
