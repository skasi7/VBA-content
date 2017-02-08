---
title: ThreeDFormat Members (Word)
ms.prod: WORD
ms.assetid: e34f22f6-7bbb-7997-d21d-9fa3da7e404b
---


# ThreeDFormat Members (Word)
Represents a shape's three-dimensional formatting.

Represents a shape's three-dimensional formatting.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[IncrementRotationHorizontal](threedformat-incrementrotationhorizontal-method-word.md)|Horizontally rotates a shape on the x-axis using the specified incrementation value.|
|[IncrementRotationVertical](threedformat-incrementrotationvertical-method-word.md)|Vertically rotates a shape on the y-axis using the specified incrementation value.|
|[IncrementRotationX](threedformat-incrementrotationx-method-word.md)|Changes the rotation of the specified shape around the x-axis by the specified number of degrees.|
|[IncrementRotationY](threedformat-incrementrotationy-method-word.md)|Changes the rotation of the specified shape around the y-axis by the specified number of degrees.|
|[IncrementRotationZ](threedformat-incrementrotationz-method-word.md)|Rotates a shape on the z-axis using the specified incrementation.|
|[ResetRotation](threedformat-resetrotation-method-word.md)|Resets the extrusion rotation around the x-axis and the y-axis to 0 (zero) so that the front of the extrusion faces forward.|
|[SetExtrusionDirection](threedformat-setextrusiondirection-method-word.md)|Sets the direction that the extrusion's sweep path takes away from the extruded shape.|
|[SetPresetCamera](threedformat-setpresetcamera-method-word.md)|Sets the camera presets for a shape.|
|[SetThreeDFormat](threedformat-setthreedformat-method-word.md)|Sets the preset extrusion format.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](threedformat-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[BevelBottomDepth](threedformat-bevelbottomdepth-property-word.md)|Returns or sets a  **Single** that represents the depth of the bottom bevel. Read/write.|
|[BevelBottomInset](threedformat-bevelbottominset-property-word.md)|Returns or sets a  **Single** that represents the inset size for the bottom bevel. Read/write.|
|[BevelBottomType](threedformat-bevelbottomtype-property-word.md)|Returns or sets an  **MsoPresetCamera** constant that represents the bevel type for the bottom bevel. Read/write.|
|[BevelTopDepth](threedformat-beveltopdepth-property-word.md)|Returns or sets a  **Single** that represents the depth of the top bevel. Read/write.|
|[BevelTopInset](threedformat-beveltopinset-property-word.md)|Returns or sets a  **Single** that represents the inset size for the top bevel. Read/write.|
|[BevelTopType](threedformat-beveltoptype-property-word.md)|Returns or sets an  **MsoPresetCamera** constant that represents the bevel type for the top bevel. Read/write.|
|[ContourColor](threedformat-contourcolor-property-word.md)|Returns or sets a  **ColorFormat** object that represents color of the contour of a shape. Read/write.|
|[ContourWidth](threedformat-contourwidth-property-word.md)|Returns or sets a  **Single** that represents the width of the contour of a shape. Read/write.|
|[Creator](threedformat-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[Depth](threedformat-depth-property-word.md)|Returns or sets the depth of the shape's extrusion. Read/write  **Single** .|
|[ExtrusionColor](threedformat-extrusioncolor-property-word.md)|Returns a  **[ColorFormat](colorformat-object-word.md)** object that represents the color of the shape's extrusion. Read-only.|
|[ExtrusionColorType](threedformat-extrusioncolortype-property-word.md)|Returns or sets a value that indicates whether the extrusion color is based on the extruded shape's fill (the front face of the extrusion) and automatically changes when the shape's fill changes, or whether the extrusion color is independent of the shape's fill. Read/write  **MsoExtrusionColorType** .|
|[FieldOfView](threedformat-fieldofview-property-word.md)|Returns or sets a  **Single** that represents the amount of perspective for a shape. Read/write.|
|[LightAngle](threedformat-lightangle-property-word.md)|Returns or sets a  **Single** that represents angle of the lighting. Read/write.|
|[Parent](threedformat-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **ThreeDFormat** object.|
|[Perspective](threedformat-perspective-property-word.md)| **MsoTrue** if the extrusion appears in perspective — that is, if the walls of the extrusion narrow toward a vanishing point. **MsoFalse** if the extrusion is a parallel, or orthographic, projection — that is, if the walls don't narrow toward a vanishing point. Read/write **MsoTriState** .|
|[PresetCamera](threedformat-presetcamera-property-word.md)|Returns an  **MsoPresetCamera** constant that represents the camera presets. Read-only.|
|[PresetExtrusionDirection](threedformat-presetextrusiondirection-property-word.md)|Returns the direction taken by the extrusion's sweep path leading away from the extruded shape (the front face of the extrusion). Read/write  **MsoPresetExtrusionDirection** .|
|[PresetLighting](threedformat-presetlighting-property-word.md)|Returns or sets an  **MsoBevelType** constant that represents the lighting preset. Read/write.|
|[PresetLightingDirection](threedformat-presetlightingdirection-property-word.md)|Returns or sets the position of the light source relative to the extrusion. Read/write  **MsoPresetLightingDirection** .|
|[PresetLightingSoftness](threedformat-presetlightingsoftness-property-word.md)|Returns or sets the intensity of the extrusion lighting. Read/write  **MsoPresetLightingSoftness** .|
|[PresetMaterial](threedformat-presetmaterial-property-word.md)|Returns or sets the extrusion surface material. Read/write  **MsoPresetMaterial** .|
|[PresetThreeDFormat](threedformat-presetthreedformat-property-word.md)|Returns the preset extrusion format. Read-only  **MsoPresetThreeDFormat** .|
|[ProjectText](threedformat-projecttext-property-word.md)|Returns or sets an  **MsoTriState** constant that represents whether text on a shape rotates with shape. **msoTrue** rotates the text. Read/write.|
|[RotationX](threedformat-rotationx-property-word.md)|Returns or sets the rotation of the extruded shape around the x-axis in degrees. Read/write  **Single** .|
|[RotationY](threedformat-rotationy-property-word.md)|Returns or sets the rotation of the extruded shape around the y-axis, in degrees. Read/write  **Single** .|
|[RotationZ](threedformat-rotationz-property-word.md)|Returns or sets a  **Single** that represents z-axis rotation of the camera. Read/write.|
|[Visible](threedformat-visible-property-word.md)| **True** if the specified object, or the formatting applied to it, is visible. Read/write **MsoTriState** .|
|[Z](threedformat-z-property-word.md)|Returns or sets a  **Single** that represents the z-axis for the shape. Read/write.|

