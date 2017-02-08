---
title: BulletFormat Object (PowerPoint)
keywords: vbapp10.chm577000
f1_keywords:
- vbapp10.chm577000
ms.prod: POWERPOINT
ms.assetid: 8c70b2af-0175-9315-3a7e-e30aa0438798
---


# BulletFormat Object (PowerPoint)

Represents bullet formatting.


## Example

Use the [Bullet](http://msdn.microsoft.com/library/paragraphformat-bullet-property-powerpoint%28Office.15%29.aspx)property to return the  **BulletFormat** object. The following example sets the bullet size and color for the paragraphs in shape two on slide one in the active presentation.


```
With ActivePresentation.Slides(1).Shapes(2)

    With .TextFrame.TextRange.ParagraphFormat.Bullet

        .Visible = True

        .RelativeSize = 1.25

        .Character = 169

        With .Font

            .Color.RGB = RGB(255, 255, 0)

            .Name = "Symbol"

        End With

    End With

End With
```


## Methods



|**Name**|
|:-----|
|[Picture](http://msdn.microsoft.com/library/bulletformat-picture-method-powerpoint%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/bulletformat-application-property-powerpoint%28Office.15%29.aspx)|
|[Character](http://msdn.microsoft.com/library/bulletformat-character-property-powerpoint%28Office.15%29.aspx)|
|[Font](http://msdn.microsoft.com/library/bulletformat-font-property-powerpoint%28Office.15%29.aspx)|
|[Number](http://msdn.microsoft.com/library/bulletformat-number-property-powerpoint%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/bulletformat-parent-property-powerpoint%28Office.15%29.aspx)|
|[RelativeSize](http://msdn.microsoft.com/library/bulletformat-relativesize-property-powerpoint%28Office.15%29.aspx)|
|[StartValue](http://msdn.microsoft.com/library/bulletformat-startvalue-property-powerpoint%28Office.15%29.aspx)|
|[Style](http://msdn.microsoft.com/library/bulletformat-style-property-powerpoint%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/bulletformat-type-property-powerpoint%28Office.15%29.aspx)|
|[UseTextColor](http://msdn.microsoft.com/library/bulletformat-usetextcolor-property-powerpoint%28Office.15%29.aspx)|
|[UseTextFont](http://msdn.microsoft.com/library/bulletformat-usetextfont-property-powerpoint%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/object-model-powerpoint-vba-reference%28Office.15%29.aspx)
