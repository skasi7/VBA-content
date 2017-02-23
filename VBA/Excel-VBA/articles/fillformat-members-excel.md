---
title: FillFormat Members (Excel)
ms.prod: EXCEL
ms.assetid: da1a1680-4b9d-c6fb-6562-bf1ec9f57921
---


# FillFormat Members (Excel)
Represents fill formatting for a shape.

Represents fill formatting for a shape.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[OneColorGradient](fillformat-onecolorgradient-method-excel.md)|Sets the specified fill to a one-color gradient.|
|[Patterned](fillformat-patterned-method-excel.md)|Sets the specified fill to a pattern.|
|[PresetGradient](fillformat-presetgradient-method-excel.md)|Sets the specified fill to a preset gradient.|
|[PresetTextured](fillformat-presettextured-method-excel.md)|Sets the specified fill format to a preset texture.|
|[Solid](fillformat-solid-method-excel.md)|Sets the specified fill to a uniform color. Use this method to convert a gradient, textured, patterned, or background fill back to a solid fill.|
|[TwoColorGradient](fillformat-twocolorgradient-method-excel.md)|Sets the specified fill to a two-color gradient.|
|[UserPicture](fillformat-userpicture-method-excel.md)|Fills the specified shape with an image.|
|[UserTextured](fillformat-usertextured-method-excel.md)|Fills the specified shape with small tiles of an image. If you want to fill the shape with one large image, use the  **[UserPicture](fillformat-userpicture-method-excel.md)** method.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](fillformat-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[BackColor](fillformat-backcolor-property-excel.md)|Returns or sets a  **[ColorFormat](colorformat-object-excel.md)** object that represents the specified fill background color.|
|[Creator](fillformat-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[ForeColor](fillformat-forecolor-property-excel.md)|Returns or sets a  **[ColorFormat](colorformat-object-excel.md)** object that represents the specified foreground fill or solid color.|
|[GradientAngle](fillformat-gradientangle-property-excel.md)|Returns or sets the angle of the gradient fill for the specified fill format. Read/write|
|[GradientColorType](fillformat-gradientcolortype-property-excel.md)|Returns the gradient color type for the specified fill. Read-only  **[MsoGradientColorType](msogradientcolortype-enumeration-office.md)** .|
|[GradientDegree](fillformat-gradientdegree-property-excel.md)|Returns the gradient degree of the specified one-color shaded fill as a floating-point value from 0.0 (dark) through 1.0 (light). Read-only  **Single** .|
|[GradientStops](fillformat-gradientstops-property-excel.md)|Returns the end point for the gradient fill. Read-only.|
|[GradientStyle](fillformat-gradientstyle-property-excel.md)|Returns the gradient style for the specified fill. Read-only  **[MsoGradientStyle](msogradientstyle-enumeration-office.md)** .|
|[GradientVariant](fillformat-gradientvariant-property-excel.md)|Returns the shade variant for the specified fill as an integer value from 1 through 4. The values for this property correspond to the gradient variants (numbered from left to right and from top to bottom) on the  **Gradient** tab in the **Fill Effects** dialog box. Read-only **Long**|
|[Parent](fillformat-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[Pattern](fillformat-pattern-property-excel.md)|Returns or sets an  **[MsoPatternType](msopatterntype-enumeration-office.md)** value that represents the fill pattern.|
|[PictureEffects](fillformat-pictureeffects-property-excel.md)|Returns an object that represents the picture or texture fill for the specified fill format. Read-only|
|[PresetGradientType](fillformat-presetgradienttype-property-excel.md)|Returns the preset gradient type for the specified fill. Read-only  **[MsoPresetGradientType](msopresetgradienttype-enumeration-office.md)** .|
|[PresetTexture](fillformat-presettexture-property-excel.md)|Returns the preset texture for the specified fill. Read-only  **[MsoPresetTexture](msopresettexture-enumeration-office.md)** .|
|[RotateWithObject](fillformat-rotatewithobject-property-excel.md)|Returns or sets if the fill style should rotate with the object. Read/write  **[MsoTriState](msotristate-enumeration-office.md)** .|
|[TextureAlignment](fillformat-texturealignment-property-excel.md)|Returns or sets the text alignment for the specified  **FillFormat** object. Read/write.|
|[TextureHorizontalScale](fillformat-texturehorizontalscale-property-excel.md)|Returns or sets the value for horizontally scaling the text for the  **FillFormat** object. Read/write **Single** .|
|[TextureName](fillformat-texturename-property-excel.md)|Returns the name of the custom texture file for the specified fill. Read-only  **String** .|
|[TextureOffsetX](fillformat-textureoffsetx-property-excel.md)|Returns the offset X value for the specified fill. Read/write  **Single** .|
|[TextureOffsetY](fillformat-textureoffsety-property-excel.md)|Returns the offset Y value for the specified fill. Read/write  **Single** .|
|[TextureTile](fillformat-texturetile-property-excel.md)|Returns the texture tile style for the specified fill. Read/write  **[MsoTriState](msotristate-enumeration-office.md)** .|
|[TextureType](fillformat-texturetype-property-excel.md)|Returns the texture type for the specified fill. Read-only  **[MsoTextureType](msotexturetype-enumeration-office.md)** .|
|[TextureVerticalScale](fillformat-textureverticalscale-property-excel.md)|Returns the texture vertical scale for the specified fill. Read/write  **Single** .|
|[Transparency](fillformat-transparency-property-excel.md)|Returns or sets the degree of transparency of the specified fill as a value from 0.0 (opaque) through 1.0 (clear). Read/write  **Double** .|
|[Type](fillformat-type-property-excel.md)|Returns a  **[MsoFillType](msofilltype-enumeration-office.md)** value that represents the fill type.|
|[Visible](fillformat-visible-property-excel.md)|Returns or sets a  **[MsoTriState](msotristate-enumeration-office.md)** value that determines whether the object is visible. Read/write.|

