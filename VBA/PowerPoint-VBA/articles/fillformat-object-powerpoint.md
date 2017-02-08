---
title: FillFormat Object (PowerPoint)
keywords: vbapp10.chm552000
f1_keywords:
- vbapp10.chm552000
ms.prod: POWERPOINT
ms.assetid: 5bd4e2cb-4466-b468-d494-bec30ed5c9d8
---


# FillFormat Object (PowerPoint)

Represents fill formatting for a shape. A shape can have a solid, gradient, texture, pattern, picture, or semi-transparent fill.


## Remarks

Many of the properties of the  **FillFormat** object are read-only. To set one of these properties, you have to apply the corresponding method.


## Example

Use the  **Fill** property to return a **FillFormat** object. The following example adds a rectangle to `myDocument` and then sets the gradient and color for the rectangle's fill.


```
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes _

        .AddShape(msoShapeRectangle, 90, 90, 90, 80).Fill

    .ForeColor.RGB = RGB(0, 128, 128)

    .OneColorGradient msoGradientHorizontal, 1, 1

End With
```


## Methods



|**Name**|
|:-----|
|[Background](http://msdn.microsoft.com/library/fillformat-background-method-powerpoint%28Office.15%29.aspx)|
|[OneColorGradient](http://msdn.microsoft.com/library/fillformat-onecolorgradient-method-powerpoint%28Office.15%29.aspx)|
|[Patterned](http://msdn.microsoft.com/library/fillformat-patterned-method-powerpoint%28Office.15%29.aspx)|
|[PresetGradient](http://msdn.microsoft.com/library/fillformat-presetgradient-method-powerpoint%28Office.15%29.aspx)|
|[PresetTextured](http://msdn.microsoft.com/library/fillformat-presettextured-method-powerpoint%28Office.15%29.aspx)|
|[Solid](http://msdn.microsoft.com/library/fillformat-solid-method-powerpoint%28Office.15%29.aspx)|
|[TwoColorGradient](http://msdn.microsoft.com/library/fillformat-twocolorgradient-method-powerpoint%28Office.15%29.aspx)|
|[UserPicture](http://msdn.microsoft.com/library/fillformat-userpicture-method-powerpoint%28Office.15%29.aspx)|
|[UserTextured](http://msdn.microsoft.com/library/fillformat-usertextured-method-powerpoint%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/fillformat-application-property-powerpoint%28Office.15%29.aspx)|
|[BackColor](http://msdn.microsoft.com/library/fillformat-backcolor-property-powerpoint%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/fillformat-creator-property-powerpoint%28Office.15%29.aspx)|
|[ForeColor](http://msdn.microsoft.com/library/fillformat-forecolor-property-powerpoint%28Office.15%29.aspx)|
|[GradientAngle](http://msdn.microsoft.com/library/fillformat-gradientangle-property-powerpoint%28Office.15%29.aspx)|
|[GradientColorType](http://msdn.microsoft.com/library/fillformat-gradientcolortype-property-powerpoint%28Office.15%29.aspx)|
|[GradientDegree](http://msdn.microsoft.com/library/fillformat-gradientdegree-property-powerpoint%28Office.15%29.aspx)|
|[GradientStops](http://msdn.microsoft.com/library/fillformat-gradientstops-property-powerpoint%28Office.15%29.aspx)|
|[GradientStyle](http://msdn.microsoft.com/library/fillformat-gradientstyle-property-powerpoint%28Office.15%29.aspx)|
|[GradientVariant](http://msdn.microsoft.com/library/fillformat-gradientvariant-property-powerpoint%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/fillformat-parent-property-powerpoint%28Office.15%29.aspx)|
|[Pattern](http://msdn.microsoft.com/library/fillformat-pattern-property-powerpoint%28Office.15%29.aspx)|
|[PictureEffects](http://msdn.microsoft.com/library/fillformat-pictureeffects-property-powerpoint%28Office.15%29.aspx)|
|[PresetGradientType](http://msdn.microsoft.com/library/fillformat-presetgradienttype-property-powerpoint%28Office.15%29.aspx)|
|[PresetTexture](http://msdn.microsoft.com/library/fillformat-presettexture-property-powerpoint%28Office.15%29.aspx)|
|[RotateWithObject](http://msdn.microsoft.com/library/fillformat-rotatewithobject-property-powerpoint%28Office.15%29.aspx)|
|[TextureAlignment](http://msdn.microsoft.com/library/fillformat-texturealignment-property-powerpoint%28Office.15%29.aspx)|
|[TextureHorizontalScale](http://msdn.microsoft.com/library/fillformat-texturehorizontalscale-property-powerpoint%28Office.15%29.aspx)|
|[TextureName](http://msdn.microsoft.com/library/fillformat-texturename-property-powerpoint%28Office.15%29.aspx)|
|[TextureOffsetX](http://msdn.microsoft.com/library/fillformat-textureoffsetx-property-powerpoint%28Office.15%29.aspx)|
|[TextureOffsetY](http://msdn.microsoft.com/library/fillformat-textureoffsety-property-powerpoint%28Office.15%29.aspx)|
|[TextureTile](http://msdn.microsoft.com/library/fillformat-texturetile-property-powerpoint%28Office.15%29.aspx)|
|[TextureType](http://msdn.microsoft.com/library/fillformat-texturetype-property-powerpoint%28Office.15%29.aspx)|
|[TextureVerticalScale](http://msdn.microsoft.com/library/fillformat-textureverticalscale-property-powerpoint%28Office.15%29.aspx)|
|[Transparency](http://msdn.microsoft.com/library/fillformat-transparency-property-powerpoint%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/fillformat-type-property-powerpoint%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/fillformat-visible-property-powerpoint%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/object-model-powerpoint-vba-reference%28Office.15%29.aspx)
