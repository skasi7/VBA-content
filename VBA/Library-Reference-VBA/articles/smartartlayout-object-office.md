---
title: SmartArtLayout Object (Office)
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.SmartArtLayout
ms.assetid: f8d9db83-86f7-4830-096d-5d15368ab6b1
---


# SmartArtLayout Object (Office)

Represents a Smart Art diagram.


## Remarks

Choices include Basic Block List, Picture Caption List, Vertical Bulleted List, etc.


## Example

The following code changes the diagram style of a Smart Art diagram in Microsoft PowerPoint.


```
ActivePresentation.Slides(1).Shapes(1).SmartArt.Layout = Application.SmartArtLayouts(1)
```


## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/smartartlayout-application-property-office%28Office.15%29.aspx)|
|[Category](http://msdn.microsoft.com/library/smartartlayout-category-property-office%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/smartartlayout-creator-property-office%28Office.15%29.aspx)|
|[Description](http://msdn.microsoft.com/library/smartartlayout-description-property-office%28Office.15%29.aspx)|
|[Id](http://msdn.microsoft.com/library/smartartlayout-id-property-office%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/smartartlayout-name-property-office%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/smartartlayout-parent-property-office%28Office.15%29.aspx)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/reference-object-library-reference-for-office%28Office.15%29.aspx)
[SmartArtLayout Object Members](http://msdn.microsoft.com/library/smartartlayout-members-office%28Office.15%29.aspx)
