---
title: PictureFormat Members (Word)
ms.prod: WORD
ms.assetid: c69a5fdb-4cd7-ee90-c58c-a423741614d6
---


# PictureFormat Members (Word)
Contains properties and methods that apply to pictures and OLE objects. The  **LinkFormat** object contains properties and methods that apply to linked OLE objects only. The **OLEFormat** object contains properties and methods that apply to OLE objects whether or not they're linked.

Contains properties and methods that apply to pictures and OLE objects. The  **LinkFormat** object contains properties and methods that apply to linked OLE objects only. The **OLEFormat** object contains properties and methods that apply to OLE objects whether or not they're linked.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[IncrementBrightness](pictureformat-incrementbrightness-method-word.md)|Changes the brightness of the picture by the specified amount.|
|[IncrementContrast](pictureformat-incrementcontrast-method-word.md)|Changes the contrast of the picture by the specified amount.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](pictureformat-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[Brightness](pictureformat-brightness-property-word.md)|Returns or sets the brightness of the specified picture or OLE object. The value for this property must be a number from 0.0 (dimmest) to 1.0 (brightest). Read/write  **Single** .|
|[ColorType](pictureformat-colortype-property-word.md)|Returns or sets the type of color transformation applied to the specified picture or OLE object. Read/write  **MsoPictureColorType** .|
|[Contrast](pictureformat-contrast-property-word.md)|Returns or sets the contrast for the specified picture or OLE object. The value for this property must be a number from 0.0 (the least contrast) to 1.0 (the greatest contrast). Read/write  **Single** .|
|[Creator](pictureformat-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[Crop](pictureformat-crop-property-word.md)|Returns or sets a [Crop](crop-object-office.md) object that represents an image cropping. Read/write.|
|[CropBottom](pictureformat-cropbottom-property-word.md)|Returns or sets the number of points that are cropped off the bottom of the specified picture or OLE object. Read/write  **Single** .|
|[CropLeft](pictureformat-cropleft-property-word.md)|Returns or sets the number of points that are cropped off the left side of the specified picture or OLE object. Read/write  **Single** .|
|[CropRight](pictureformat-cropright-property-word.md)|Returns or sets the number of points that are cropped off the right side of the specified picture or OLE object. Read/write  **Single** .|
|[CropTop](pictureformat-croptop-property-word.md)|Returns or sets the number of points that are cropped off the top of the specified picture or OLE object. Read/write  **Single** .|
|[Parent](pictureformat-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **PictureFormat** object.|
|[TransparencyColor](pictureformat-transparencycolor-property-word.md)|Returns or sets the transparent color for the specified picture as a red-green-blue (RGB) value. Read/write  **Long** .|
|[TransparentBackground](pictureformat-transparentbackground-property-word.md)| **MsoTrue** if the parts of the picture that are defined with a transparent color actually appear transparent. Use the **TransparencyColor** property to set the transparent color. Applies to bitmaps only. Read/write **MsoTriState** .|

