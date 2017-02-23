---
title: PictureFormat Members (Excel)
ms.prod: EXCEL
ms.assetid: d27d6074-2698-2b1d-87cb-c9cc187354c3
---


# PictureFormat Members (Excel)
Contains properties and methods that apply to pictures and OLE objects.

Contains properties and methods that apply to pictures and OLE objects.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[IncrementBrightness](pictureformat-incrementbrightness-method-excel.md)|Changes the brightness of the picture by the specified amount. Use the  **[Brightness](pictureformat-brightness-property-excel.md)** property to set the absolute brightness of the picture.|
|[IncrementContrast](pictureformat-incrementcontrast-method-excel.md)|Changes the contrast of the picture by the specified amount. Use the  **[Contrast](pictureformat-contrast-property-excel.md)** property to set the absolute contrast for the picture.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](pictureformat-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[Brightness](pictureformat-brightness-property-excel.md)|Returns or sets the brightness of the specified picture or OLE object. The value for this property must be a number from 0.0 (dimmest) to 1.0 (brightest). Read/write  **Single** .|
|[ColorType](pictureformat-colortype-property-excel.md)|Returns or sets the type of color transformation applied to the specified picture or OLE object. Read/write .|
|[Contrast](pictureformat-contrast-property-excel.md)|Returns or sets the contrast for the specified picture or OLE object. The value for this property must be a number from 0.0 (the least contrast) to 1.0 (the greatest contrast). Read/write  **Single** .|
|[Creator](pictureformat-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[Crop](pictureformat-crop-property-excel.md)|Returns an  **[Crop](crop-object-office.md)** object that represents the cropping settings for the specified **[PictureFormat](pictureformat-object-excel.md)** object. Read-only|
|[CropBottom](pictureformat-cropbottom-property-excel.md)|Returns or sets the number of points that are cropped off the bottom of the specified picture or OLE object. Read/write  **Single** .|
|[CropLeft](pictureformat-cropleft-property-excel.md)|Returns or sets the number of points that are cropped off the left side of the specified picture or OLE object. Read/write  **Single** .|
|[CropRight](pictureformat-cropright-property-excel.md)|Returns or sets the number of points that are cropped off the right side of the specified picture or OLE object. Read/write  **Single** .|
|[CropTop](pictureformat-croptop-property-excel.md)|Returns or sets the number of points that are cropped off the top of the specified picture or OLE object. Read/write  **Single** .|
|[Parent](pictureformat-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[TransparencyColor](pictureformat-transparencycolor-property-excel.md)|Returns or sets the transparent color for the specified picture as a red-green-blue (RGB) value. For this property to take effect, the  **[TransparentBackground](pictureformat-transparentbackground-property-excel.md)** property must be set to **True** . Applies to bitmaps only. Read/write **Long** .|
|[TransparentBackground](pictureformat-transparentbackground-property-excel.md)|Use the  **[TransparencyColor](pictureformat-transparencycolor-property-excel.md)** property to set the transparent color. Applies to bitmaps only. Read/write MsoTriState.|

