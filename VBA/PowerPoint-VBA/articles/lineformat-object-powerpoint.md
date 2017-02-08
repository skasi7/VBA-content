---
title: LineFormat Object (PowerPoint)
keywords: vbapp10.chm553000
f1_keywords:
- vbapp10.chm553000
ms.prod: POWERPOINT
ms.assetid: 11c955d5-bbda-d99f-cec9-fc6187450a12
---


# LineFormat Object (PowerPoint)

Represents line and arrowhead formatting. For a line, the  **LineFormat** object contains formatting information for the line itself; for a shape with a border, this object contains formatting information for the shape's border.


## Example

Use the  **Line** property to return a **LineFormat** object. The following example adds a blue, dashed line to `myDocument`. There's a short, narrow oval at the line's starting point and a long, wide triangle at its endpoint.


```
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes.AddLine(100, 100, 200, 300).Line

    .DashStyle = msoLineDashDotDot

    .ForeColor.RGB = RGB(50, 0, 128)

    .BeginArrowheadLength = msoArrowheadShort

    .BeginArrowheadStyle = msoArrowheadOval

    .BeginArrowheadWidth = msoArrowheadNarrow

    .EndArrowheadLength = msoArrowheadLong

    .EndArrowheadStyle = msoArrowheadTriangle

    .EndArrowheadWidth = msoArrowheadWide

End With
```


## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/lineformat-application-property-powerpoint%28Office.15%29.aspx)|
|[BackColor](http://msdn.microsoft.com/library/lineformat-backcolor-property-powerpoint%28Office.15%29.aspx)|
|[BeginArrowheadLength](http://msdn.microsoft.com/library/lineformat-beginarrowheadlength-property-powerpoint%28Office.15%29.aspx)|
|[BeginArrowheadStyle](http://msdn.microsoft.com/library/lineformat-beginarrowheadstyle-property-powerpoint%28Office.15%29.aspx)|
|[BeginArrowheadWidth](http://msdn.microsoft.com/library/lineformat-beginarrowheadwidth-property-powerpoint%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/lineformat-creator-property-powerpoint%28Office.15%29.aspx)|
|[DashStyle](http://msdn.microsoft.com/library/lineformat-dashstyle-property-powerpoint%28Office.15%29.aspx)|
|[EndArrowheadLength](http://msdn.microsoft.com/library/lineformat-endarrowheadlength-property-powerpoint%28Office.15%29.aspx)|
|[EndArrowheadStyle](http://msdn.microsoft.com/library/lineformat-endarrowheadstyle-property-powerpoint%28Office.15%29.aspx)|
|[EndArrowheadWidth](http://msdn.microsoft.com/library/lineformat-endarrowheadwidth-property-powerpoint%28Office.15%29.aspx)|
|[ForeColor](http://msdn.microsoft.com/library/lineformat-forecolor-property-powerpoint%28Office.15%29.aspx)|
|[InsetPen](http://msdn.microsoft.com/library/lineformat-insetpen-property-powerpoint%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/lineformat-parent-property-powerpoint%28Office.15%29.aspx)|
|[Pattern](http://msdn.microsoft.com/library/lineformat-pattern-property-powerpoint%28Office.15%29.aspx)|
|[Style](http://msdn.microsoft.com/library/lineformat-style-property-powerpoint%28Office.15%29.aspx)|
|[Transparency](http://msdn.microsoft.com/library/lineformat-transparency-property-powerpoint%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/lineformat-visible-property-powerpoint%28Office.15%29.aspx)|
|[Weight](http://msdn.microsoft.com/library/lineformat-weight-property-powerpoint%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/object-model-powerpoint-vba-reference%28Office.15%29.aspx)
