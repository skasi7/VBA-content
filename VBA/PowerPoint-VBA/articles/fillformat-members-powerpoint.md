---
title: FillFormat Members (PowerPoint)
ms.prod: POWERPOINT
ms.assetid: ccd26632-4ff8-6fad-2c5d-c26078eeff3b
---


# FillFormat Members (PowerPoint)
Represents fill formatting for a shape. A shape can have a solid, gradient, texture, pattern, picture, or semi-transparent fill.

Represents fill formatting for a shape. A shape can have a solid, gradient, texture, pattern, picture, or semi-transparent fill.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Background](fillformat-background-method-powerpoint.md)|Specifies that the shape's fill should match the slide background. If you change the slide background after applying this method to a fill, the fill will also change.|
|[OneColorGradient](fillformat-onecolorgradient-method-powerpoint.md)|Sets the specified fill to a one-color gradient.|
|[Patterned](fillformat-patterned-method-powerpoint.md)|Sets the specified fill to a pattern.|
|[PresetGradient](fillformat-presetgradient-method-powerpoint.md)|Sets the specified fill to a preset gradient.|
|[PresetTextured](fillformat-presettextured-method-powerpoint.md)|Sets the specified fill to a preset texture.|
|[Solid](fillformat-solid-method-powerpoint.md)|Sets the specified fill to a uniform color. Use this method to convert a gradient, textured, patterned, or background fill back to a solid fill.|
|[TwoColorGradient](fillformat-twocolorgradient-method-powerpoint.md)|Sets the specified fill to a two-color gradient.|
|[UserPicture](fillformat-userpicture-method-powerpoint.md)|Fills the specified shape with one large image. |
|[UserTextured](fillformat-usertextured-method-powerpoint.md)|Fills the specified shape with small tiles of an image. |

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](fillformat-application-property-powerpoint.md)|Returns an  **[Application](application-object-powerpoint.md)** object that represents the creator of the specified object.|
|[BackColor](fillformat-backcolor-property-powerpoint.md)|Returns or sets a  **[ColorFormat](colorformat-object-powerpoint.md)** object that represents the background color for the specified fill or patterned line. Read/write.|
|[Creator](fillformat-creator-property-powerpoint.md)|Returns a  **Long** that represents the four-character creator code for the application in which the specified object was created. For example, if the object was created in Microsoft PowerPoint, this property returns the hexadecimal number 50575054. Read-only.|
|[ForeColor](fillformat-forecolor-property-powerpoint.md)|Returns or sets a  **[ColorFormat](colorformat-object-powerpoint.md)** object that represents the foreground color for the fill, line, or shadow. Read/write.|
|[GradientAngle](fillformat-gradientangle-property-powerpoint.md)|Returns or sets the angle of the gradient fill for the specified fill format. Read/write.|
|[GradientColorType](fillformat-gradientcolortype-property-powerpoint.md)|Returns the gradient color type for the specified fill. Read-only.|
|[GradientDegree](fillformat-gradientdegree-property-powerpoint.md)|Returns a value that indicates how dark or light a one-color gradient fill is. Read-only.|
|[GradientStops](fillformat-gradientstops-property-powerpoint.md)| Returns the **[GradientStops](gradientstops-object-office.md)** collection associated with the specified fill format. Read-only.|
|[GradientStyle](fillformat-gradientstyle-property-powerpoint.md)|Returns the gradient style for the specified fill. Read-only.|
|[GradientVariant](fillformat-gradientvariant-property-powerpoint.md)|Returns the gradient variant for the specified fill as an integer value from 1 to 4 for most gradient fills. Read-only.|
|[Parent](fillformat-parent-property-powerpoint.md)|Returns the parent object for the specified object.|
|[Pattern](fillformat-pattern-property-powerpoint.md)|Sets or returns a value that represents the pattern applied to the specified fill. Read-only.|
|[PictureEffects](fillformat-pictureeffects-property-powerpoint.md)|Returns an object that represents the picture or texture fill for the specified fill format. Read-only.|
|[PresetGradientType](fillformat-presetgradienttype-property-powerpoint.md)|Returns the preset gradient type for the specified fill. Read-only. |
|[PresetTexture](fillformat-presettexture-property-powerpoint.md)|Returns the preset texture for the specified fill. Read-only.|
|[RotateWithObject](fillformat-rotatewithobject-property-powerpoint.md)|Returns or sets whether the fill rotates with the specified shape. Read/write.|
|[TextureAlignment](fillformat-texturealignment-property-powerpoint.md)|Returns or sets the alignment (the origin of the coordinate grid) for the tiling of the texture fill. Read/write.|
|[TextureHorizontalScale](fillformat-texturehorizontalscale-property-powerpoint.md)|Returns or sets the horizontal scaling factor for the texture fill. Read/write.|
|[TextureName](fillformat-texturename-property-powerpoint.md)|Returns the name of the custom texture file for the specified fill. Read-only.|
|[TextureOffsetX](fillformat-textureoffsetx-property-powerpoint.md)| Returns or sets the horizontal offset of the texture from the origin in points. Read/write.|
|[TextureOffsetY](fillformat-textureoffsety-property-powerpoint.md)| Returns or sets the vertical offset of the texture from the origin in points. Read/write.|
|[TextureTile](fillformat-texturetile-property-powerpoint.md)| Returns or sets whether the texture fill is tiled or centered. Read/write.|
|[TextureType](fillformat-texturetype-property-powerpoint.md)|Returns the texture type for the specified fill. Read-only.|
|[TextureVerticalScale](fillformat-textureverticalscale-property-powerpoint.md)|Returns or sets the vertical scaling factor for the texture fill. Read/write.|
|[Transparency](fillformat-transparency-property-powerpoint.md)|Returns or sets the degree of transparency of the specified fill, shadow, or line as a value between 0.0 (opaque) and 1.0 (clear). Read/write.|
|[Type](fillformat-type-property-powerpoint.md)|Represent the type of fill. Read-only.|
|[Visible](fillformat-visible-property-powerpoint.md)|Returns or sets the visibility of the specified object or the formatting applied to the specified object. Read/write.|

