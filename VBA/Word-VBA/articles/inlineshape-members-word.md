---
title: InlineShape Members (Word)
ms.prod: WORD
ms.assetid: f9de7adf-d761-3824-ba2e-c58c26de3d82
---


# InlineShape Members (Word)
Represents an object in the text layer of a document. An inline shape can only be a picture, an OLE object, or an ActiveX control. The  **InlineShape** object is a member of the **[InlineShapes](inlineshapes-object-word.md)** collection. The **InlineShapes** collection contains all the shapes that appear inline in a document, range, or selection.

Represents an object in the text layer of a document. An inline shape can only be a picture, an OLE object, or an ActiveX control. The  **InlineShape** object is a member of the **[InlineShapes](inlineshapes-object-word.md)** collection. The **InlineShapes** collection contains all the shapes that appear inline in a document, range, or selection.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[ConvertToShape](inlineshape-converttoshape-method-word.md)|Converts an inline shape to a free-floating shape. Returns a  **[Shape](shape-object-word.md)** object that represents the new shape.|
|[Delete](inlineshape-delete-method-word.md)|Deletes the specified inline shape.|
|[Reset](inlineshape-reset-method-word.md)|Removes changes that were made to an inline shape.|
|[Select](inlineshape-select-method-word.md)|Selects the specified inline shape.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[AlternativeText](inlineshape-alternativetext-property-word.md)|Returns or sets a  **String** that represents the alternative text associated with a shape in a Web page. Read/write.|
|[Application](inlineshape-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[Borders](inlineshape-borders-property-word.md)|Returns a  **[Borders](borders-object-word.md)** collection that represents all the borders for the specified shape.|
|[Chart](inlineshape-chart-property-word.md)|Returns a  **Chart** object that represents a chart within the collection of inline shapes in a document. Read-only.|
|[Creator](inlineshape-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[Field](inlineshape-field-property-word.md)|Returns a  **Field** object that represents the field associated with the specified inline shape. Read-only.|
|[Fill](inlineshape-fill-property-word.md)|Returns a  **[FillFormat](fillformat-object-word.md)** object that contains fill formatting properties for the specified shape. Read-only.|
|[Glow](inlineshape-glow-property-word.md)|Returns a  **[GlowFormat](glowformat-object-word.md)** object that represents the formatting properties for a glow effect. Read-only.|
|[GroupItems](inlineshape-groupitems-property-word.md)|Returns a  **[GroupShapes](groupshapes-object-word.md)** collection that represents the shapes that are grouped together for an inline shape. Read-only.|
|[HasChart](inlineshape-haschart-property-word.md)| **True** if the specified shape is a chart. Read-only.|
|[HasSmartArt](inlineshape-hassmartart-property-word.md)|Returns  **True** if there is a SmartArt diagram present on the shape. Read-only.|
|[Height](inlineshape-height-property-word.md)|Returns or sets the height of an inline shape. Read/write  **Single** .|
|[HorizontalLineFormat](inlineshape-horizontallineformat-property-word.md)|Returns a  **[HorizontalLineFormat](horizontallineformat-object-word.md)** object that contains the horizontal line formatting for the specified **InlineShape** object. Read-only.|
|[Hyperlink](inlineshape-hyperlink-property-word.md)|Returns a  **Hyperlink** object that represents the hyperlink associated with the specified inline shape. Read-only.|
|[IsPictureBullet](inlineshape-ispicturebullet-property-word.md)| **True** indicates that an **InlineShape** object is a picture bullet. Read-only **Boolean** .|
|[Line](inlineshape-line-property-word.md)|Returns a  **LineFormat** object that contains line formatting properties for the specified shape. Read-only.|
|[LinkFormat](inlineshape-linkformat-property-word.md)|Returns a  **LinkFormat** object that represents the link options of the specified inline shape that is linked to a file. Read/only.|
|[LockAspectRatio](inlineshape-lockaspectratio-property-word.md)| **MsoTrue** if the specified shape retains its original proportions when you resize it. **MsoFalse** if you can change the height and width of the shape independently of one another when you resize it. Read/write **MsoTriState** .|
|[OLEFormat](inlineshape-oleformat-property-word.md)|Returns an  **OLEFormat** object that represents the OLE characteristics (other than linking) for the specified inline shape. Read-only.|
|[Parent](inlineshape-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **InlineShape** object.|
|[PictureFormat](inlineshape-pictureformat-property-word.md)|Returns a  **PictureFormat** object that contains picture formatting properties for the inline shape. Read-only.|
|[Range](inlineshape-range-property-word.md)|Returns a  **Range** object that represents the portion of a document that is contained within an inline shape.|
|[Reflection](inlineshape-reflection-property-word.md)|Returns a  **[ReflectionFormat](reflectionformat-object-word.md)** object that represents the reflection formatting for a shape. Read-only.|
|[ScaleHeight](inlineshape-scaleheight-property-word.md)|Scales the height of the specified inline shape relative to its original size. Read/write  **Single** .|
|[ScaleWidth](inlineshape-scalewidth-property-word.md)|Scales the width of the specified inline shape relative to its original size. Read/write  **Single** .|
|[Script](inlineshape-script-property-word.md)|Returns a  **Script** object, which represents a block of script or code associated with an image on the specified Web page.|
|[Shadow](inlineshape-shadow-property-word.md)|Returns a  **ShadowFormat** object that represents the shadow formatting for the specified shape. Read-only.|
|[SmartArt](inlineshape-smartart-property-word.md)|Returns a [SmartArt](smartart-object-office.md) object that provides a way to work with the SmartArt associated with the specified inline shape. Read-only.|
|[SoftEdge](inlineshape-softedge-property-word.md)|Returns a  **[SoftEdgeFormat](softedgeformat-object-word.md)** object that represents the soft edge formatting for a shape. Read-only.|
|[TextEffect](inlineshape-texteffect-property-word.md)|Returns a  **TextEffectFormat** object that contains text-effect formatting properties for the specified inline shape. Read-only.|
|[Title](inlineshape-title-property-word.md)|Returns or sets a  **String** that contains a title for the specified inline shape. Read/write.|
|[Type](inlineshape-type-property-word.md)|Returns the type of inline shape. Read-only  **[WdInlineShapeType](wdinlineshapetype-enumeration-word.md)** .|
|[Width](inlineshape-width-property-word.md)|Returns or sets the width, in points, of the specified inline shape. Read/write  **Long** .|

