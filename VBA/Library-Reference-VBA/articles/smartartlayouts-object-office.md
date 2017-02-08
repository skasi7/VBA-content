---
title: SmartArtLayouts Object (Office)
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.SmartArtLayouts
ms.assetid: 25e33439-fb5e-01d7-1b85-01884a42ba68
---


# SmartArtLayouts Object (Office)

Represents a collection of Smart Art layout diagrams.


## Remarks

Choices include Basic Block List, Picture Caption List, Vertical Bulleted List, etc.


## Example

The following code changes the diagram style of a Smart Art diagram in Microsoft PowerPoint.


```
ActivePresentation.Slides(1).Shapes(1).SmartArt.Layout = Application.SmartArtLayouts(1)
```


## Methods



|**Name**|
|:-----|
|[Item](http://msdn.microsoft.com/library/smartartlayouts-item-method-office%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/smartartlayouts-application-property-office%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/smartartlayouts-count-property-office%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/smartartlayouts-creator-property-office%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/smartartlayouts-parent-property-office%28Office.15%29.aspx)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/reference-object-library-reference-for-office%28Office.15%29.aspx)
[SmartArtLayouts Object Members](http://msdn.microsoft.com/library/smartartlayouts-members-office%28Office.15%29.aspx)
