---
title: GradientStops Object (Office)
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.GradientStops
ms.assetid: 365949f0-29b3-76e1-1163-2ac870f68f7a
---


# GradientStops Object (Office)

Contains a collection of  **GradientStop** objects.


## Remarks

Gradients are a smooth transition from one color state to another. The endpoints of these sections are called stops.


## Example

The following example creates three color gradient stops in Microsoft PowerPoint.


```
Sub gradients() 
 Set myDocument = ActivePresentation.Slides(1) 
 Set GradientShapeFill = myDocument.Shapes.AddShape(msoShapeRectangle, 90, 90, 90, 80).Fill 
 With GradientShapeFill 
 .ForeColor.RGB = RGB(0, 128, 128) 
 .OneColorGradient msoGradientHorizontal, 1, 1 
 .GradientStops.Insert RGB(255, 0, 0), 0.25 
 .GradientStops.Insert RGB(0, 255, 0), 0.5 
 .GradientStops.Insert RGB(0, 0, 255), 0.75 
 End With 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Delete](http://msdn.microsoft.com/library/gradientstops-delete-method-office%28Office.15%29.aspx)|
|[Insert](http://msdn.microsoft.com/library/gradientstops-insert-method-office%28Office.15%29.aspx)|
|[Insert2](http://msdn.microsoft.com/library/gradientstops-insert2-method-office%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/gradientstops-application-property-office%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/gradientstops-count-property-office%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/gradientstops-creator-property-office%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/gradientstops-item-property-office%28Office.15%29.aspx)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/reference-object-library-reference-for-office%28Office.15%29.aspx)
[GradientStops Object Members](http://msdn.microsoft.com/library/gradientstops-members-office%28Office.15%29.aspx)
