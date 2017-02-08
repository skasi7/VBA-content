---
title: Selection Object (PowerPoint)
keywords: vbapp10.chm508000
f1_keywords:
- vbapp10.chm508000
ms.prod: POWERPOINT
ms.assetid: a7def3bd-9dff-da53-152d-4fd686642413
---


# Selection Object (PowerPoint)

Represents the selection in the specified document window. The  **Selection** object is deleted whenever you change slides in an active slide view (the **Type** property will return **ppSelectionNone** ).


## Example

Use the [Selection](http://msdn.microsoft.com/library/cell-selected-property-powerpoint%28Office.15%29.aspx)property to return the  **Selection** object. The following example places a copy of the selection in the active window on the Clipboard.


```
ActiveWindow.Selection.Copy
```

Use the [ShapeRange](http://msdn.microsoft.com/library/selection-shaperange-property-powerpoint%28Office.15%29.aspx), [SlideRange](http://msdn.microsoft.com/library/selection-sliderange-property-powerpoint%28Office.15%29.aspx), or [TextRange](http://msdn.microsoft.com/library/selection-textrange-property-powerpoint%28Office.15%29.aspx)property to return a range of shapes, slides, or text from the selection.

The following example sets the fill foreground color for the selected shapes in window two, assuming that there's at least one shape selected, and assuming that all selected shapes have a fill whose forecolor can be set.




```
With Windows(2).Selection.ShapeRange.Fill

    .Visible = True

    .ForeColor.RGB = RGB(255, 0, 255)

End With
```

The following example sets the text in the first selected shape in window two if that shape contains a text frame.




```
With Windows(2).Selection.ShapeRange(1)

    If .HasTextFrame Then

        .TextFrame.TextRange = "Current Choice"

    End If

End With
```

The following example cuts the selected text in the active window and places it on the Clipboard.




```
ActiveWindow.Selection.TextRange.Cut
```

The following example duplicates all the slides in the selection (if you're in slide view, this duplicates the current slide).




```
ActiveWindow.Selection.SlideRange.Duplicate
```

If you don't have an object of the appropriate type selected when you use one of these properties (for instance, if you use the  **ShapeRange** property when there are no shapes selected), an error occurs. Use the[Type](http://msdn.microsoft.com/library/selection-type-property-powerpoint%28Office.15%29.aspx)property to determine what kind of object or objects are selected. The following example checks to see whether the selection contains slides. If the selection does contain slides, the example sets the background for the first slide in the selection.




```
With Windows(2).Selection

    If .Type = ppSelectionSlides Then

        With .SlideRange(1)

            .FollowMasterBackground = False

            .Background.Fill.PresetGradient _

                msoGradientHorizontal, 1, msoGradientLateSunset

        End With

    End If

End With
```


## Methods



|**Name**|
|:-----|
|[Copy](http://msdn.microsoft.com/library/selection-copy-method-powerpoint%28Office.15%29.aspx)|
|[Cut](http://msdn.microsoft.com/library/selection-cut-method-powerpoint%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/selection-delete-method-powerpoint%28Office.15%29.aspx)|
|[Unselect](http://msdn.microsoft.com/library/selection-unselect-method-powerpoint%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/selection-application-property-powerpoint%28Office.15%29.aspx)|
|[ChildShapeRange](http://msdn.microsoft.com/library/selection-childshaperange-property-powerpoint%28Office.15%29.aspx)|
|[HasChildShapeRange](http://msdn.microsoft.com/library/selection-haschildshaperange-property-powerpoint%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/selection-parent-property-powerpoint%28Office.15%29.aspx)|
|[ShapeRange](http://msdn.microsoft.com/library/selection-shaperange-property-powerpoint%28Office.15%29.aspx)|
|[SlideRange](http://msdn.microsoft.com/library/selection-sliderange-property-powerpoint%28Office.15%29.aspx)|
|[TextRange](http://msdn.microsoft.com/library/selection-textrange-property-powerpoint%28Office.15%29.aspx)|
|[TextRange2](http://msdn.microsoft.com/library/selection-textrange2-property-powerpoint%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/selection-type-property-powerpoint%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/object-model-powerpoint-vba-reference%28Office.15%29.aspx)
