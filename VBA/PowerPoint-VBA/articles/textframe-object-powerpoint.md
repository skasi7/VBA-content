---
title: TextFrame Object (PowerPoint)
keywords: vbapp10.chm558000
f1_keywords:
- vbapp10.chm558000
ms.prod: POWERPOINT
ms.assetid: 03346e81-71b2-0b9e-843d-fb8aa0e3c868
---


# TextFrame Object (PowerPoint)

Represents the text frame in a  **Shape** object. Contains the text in the text frame and the properties and methods that control the alignment and anchoring of the text frame.


## Example

Use the  **TextFrame** property to return a **TextFrame** object. The following example adds a rectangle to myDocument, adds text to the rectangle, and then sets the margins for the text frame.


```
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes _

        .AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame

    .TextRange.Text = "Here is some test text"

    .MarginBottom = 10

    .MarginLeft = 10

    .MarginRight = 10

    .MarginTop = 10

End With
```

Use the [HasTextFrame](http://msdn.microsoft.com/library/shape-hastextframe-property-powerpoint%28Office.15%29.aspx)property to determine whether a shape has a text frame, and use the [HasText](http://msdn.microsoft.com/library/textframe-hastext-property-powerpoint%28Office.15%29.aspx)property to determine whether the text frame contains text, as shown in the following example.




```
Set myDocument = ActivePresentation.Slides(1)

For Each s In myDocument.Shapes

    If s.HasTextFrame Then

        With s.TextFrame

            If .HasText Then MsgBox .TextRange.Text

        End With

    End If

Next
```


## Methods



|**Name**|
|:-----|
|[DeleteText](http://msdn.microsoft.com/library/textframe-deletetext-method-powerpoint%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/textframe-application-property-powerpoint%28Office.15%29.aspx)|
|[AutoSize](http://msdn.microsoft.com/library/textframe-autosize-property-powerpoint%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/textframe-creator-property-powerpoint%28Office.15%29.aspx)|
|[HasText](http://msdn.microsoft.com/library/textframe-hastext-property-powerpoint%28Office.15%29.aspx)|
|[HorizontalAnchor](http://msdn.microsoft.com/library/textframe-horizontalanchor-property-powerpoint%28Office.15%29.aspx)|
|[MarginBottom](http://msdn.microsoft.com/library/textframe-marginbottom-property-powerpoint%28Office.15%29.aspx)|
|[MarginLeft](http://msdn.microsoft.com/library/textframe-marginleft-property-powerpoint%28Office.15%29.aspx)|
|[MarginRight](http://msdn.microsoft.com/library/textframe-marginright-property-powerpoint%28Office.15%29.aspx)|
|[MarginTop](http://msdn.microsoft.com/library/textframe-margintop-property-powerpoint%28Office.15%29.aspx)|
|[Orientation](http://msdn.microsoft.com/library/textframe-orientation-property-powerpoint%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/textframe-parent-property-powerpoint%28Office.15%29.aspx)|
|[Ruler](http://msdn.microsoft.com/library/textframe-ruler-property-powerpoint%28Office.15%29.aspx)|
|[TextRange](http://msdn.microsoft.com/library/textframe-textrange-property-powerpoint%28Office.15%29.aspx)|
|[VerticalAnchor](http://msdn.microsoft.com/library/textframe-verticalanchor-property-powerpoint%28Office.15%29.aspx)|
|[WordWrap](http://msdn.microsoft.com/library/textframe-wordwrap-property-powerpoint%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/object-model-powerpoint-vba-reference%28Office.15%29.aspx)
