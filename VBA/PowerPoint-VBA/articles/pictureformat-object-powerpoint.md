---
title: PictureFormat Object (PowerPoint)
keywords: vbapp10.chm551000
f1_keywords:
- vbapp10.chm551000
ms.prod: POWERPOINT
ms.assetid: 946794b4-0401-ec7c-cea3-779ebfce0d69
---


# PictureFormat Object (PowerPoint)

Contains properties and methods that apply to pictures and OLE objects. 


## Example

Use the  **PictureFormat** property to return a **PictureFormat** object. The following example sets the brightness, contrast, and color transformation for shape one on `myDocument` and crops 18 points off the bottom of the shape. For this example to work, shape one must be either a picture or an OLE object.


```
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(1).PictureFormat

    .Brightness = 0.3

    .Contrast = 0.7

    .ColorType = msoPictureGrayScale

    .CropBottom = 18

End With
```


## Methods



|**Name**|
|:-----|
|[IncrementBrightness](http://msdn.microsoft.com/library/pictureformat-incrementbrightness-method-powerpoint%28Office.15%29.aspx)|
|[IncrementContrast](http://msdn.microsoft.com/library/pictureformat-incrementcontrast-method-powerpoint%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/pictureformat-application-property-powerpoint%28Office.15%29.aspx)|
|[Brightness](http://msdn.microsoft.com/library/pictureformat-brightness-property-powerpoint%28Office.15%29.aspx)|
|[ColorType](http://msdn.microsoft.com/library/pictureformat-colortype-property-powerpoint%28Office.15%29.aspx)|
|[Contrast](http://msdn.microsoft.com/library/pictureformat-contrast-property-powerpoint%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/pictureformat-creator-property-powerpoint%28Office.15%29.aspx)|
|[Crop](http://msdn.microsoft.com/library/pictureformat-crop-property-powerpoint%28Office.15%29.aspx)|
|[CropBottom](http://msdn.microsoft.com/library/pictureformat-cropbottom-property-powerpoint%28Office.15%29.aspx)|
|[CropLeft](http://msdn.microsoft.com/library/pictureformat-cropleft-property-powerpoint%28Office.15%29.aspx)|
|[CropRight](http://msdn.microsoft.com/library/pictureformat-cropright-property-powerpoint%28Office.15%29.aspx)|
|[CropTop](http://msdn.microsoft.com/library/pictureformat-croptop-property-powerpoint%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/pictureformat-parent-property-powerpoint%28Office.15%29.aspx)|
|[TransparencyColor](http://msdn.microsoft.com/library/pictureformat-transparencycolor-property-powerpoint%28Office.15%29.aspx)|
|[TransparentBackground](http://msdn.microsoft.com/library/pictureformat-transparentbackground-property-powerpoint%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/object-model-powerpoint-vba-reference%28Office.15%29.aspx)
