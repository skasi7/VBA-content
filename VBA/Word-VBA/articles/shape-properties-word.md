---
title: Shape Properties (Word)
ms.prod: WORD
ms.assetid: 6d721b07-3397-4caf-afd5-730f86a1027a
---


# Shape Properties (Word)

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Adjustments](shape-adjustments-property-word.md)|Returns an  **[Adjustments](adjustments-object-word.md)** object that contains adjustment values for all the adjustments in the specified **Shape** object that represents an AutoShape or WordArt. Read-only.|
|[AlternativeText](shape-alternativetext-property-word.md)|Returns or sets the alternative text associated with a shape in a Web page. Read/write  **String** .|
|[Anchor](shape-anchor-property-word.md)|Returns a  **Range** object that represents the anchoring range for the specified shape or shape range. Read-only.|
|[Application](shape-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[AutoShapeType](shape-autoshapetype-property-word.md)|Returns or sets the shape type for the specified  **Shape** object, which must represent an AutoShape other than a line or freeform drawing. Read/write **MsoAutoShapeType** .|
|[BackgroundStyle](shape-backgroundstyle-property-word.md)|Sets or returns the background style of the specified shape. Read/write [MsoBackgroundStyleIndex](msobackgroundstyleindex-enumeration-office.md).|
|[Callout](shape-callout-property-word.md)|Returns a  **[CalloutFormat](calloutformat-object-word.md)** object that contains callout formatting properties for the specified shape. Read-only.|
|[CanvasItems](shape-canvasitems-property-word.md)|Returns a  **[CanvasShapes](canvasshapes-object-word.md)** object that represents a collection of shapes in a drawing canvas.|
|[Chart](shape-chart-property-word.md)|Returns a  **Chart** object that represents a chart within the collection of shapes in a document. Read-only.|
|[Child](shape-child-property-word.md)| **True** if the shape is a child shape or if all shapes in a shape range are child shapes of the same parent. Read-only **MsoTriState** .|
|[Creator](shape-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[Fill](shape-fill-property-word.md)|Returns a  **[FillFormat](fillformat-object-word.md)** object that contains fill formatting properties for the specified shape. Read-only.|
|[Glow](shape-glow-property-word.md)|Returns a  **[GlowFormat](glowformat-object-word.md)** object that represents the glow formatting for a shape. Read-only.|
|[GroupItems](shape-groupitems-property-word.md)|Returns a  **[GroupShapes](groupshapes-object-word.md)** object that represents the individual shapes in the specified group. Read-only.|
|[HasChart](shape-haschart-property-word.md)| **True** if the specified shape has a chart. Read-only.|
|[HasSmartArt](shape-hassmartart-property-word.md)|Returns  **True** if there is a SmartArt diagram present on the shape. Read-only.|
|[Height](shape-height-property-word.md)|Returns or sets the height of the specified shape. Read/write  **Single** .|
|[HeightRelative](shape-heightrelative-property-word.md)|Returns or sets a  **Single** that represents the percentage of the relative height of a shape. Read/write.|
|[HorizontalFlip](shape-horizontalflip-property-word.md)|Indicates that a shape has been flipped horizontally. Read-only  **MsoTriState** .|
|[Hyperlink](shape-hyperlink-property-word.md)|Returns a  **Hyperlink** object that represents the hyperlink associated with a **Shape** object. Read-only.|
|[ID](shape-id-property-word.md)|Returns the identification type for the specified shape. Read-only  **Long** .|
|[LayoutInCell](shape-layoutincell-property-word.md)|Returns a  **Long** that represents whether a shape in a table is displayed inside or outside the table.|
|[Left](shape-left-property-word.md)|Returns or sets a  **Single** that represents the horizontal position, measured in points, of the specified shape or shape range. Can also be any valid **[WdShapePosition](wdshapeposition-enumeration-word.md)** constant. Read/write.|
|[LeftRelative](shape-leftrelative-property-word.md)|Returns or sets a  **Single** that represents the relative left position of a shape. Read/write.|
|[Line](shape-line-property-word.md)|Returns a  **LineFormat** object that contains line formatting properties for the specified shape. Read-only.|
|[LinkFormat](shape-linkformat-property-word.md)|Returns a  **LinkFormat** object that represents the link options of a shape that is linked to a file. Read/only.|
|[LockAnchor](shape-lockanchor-property-word.md)| **True** if the anchor of a **Shape** object is locked to the anchoring range. Read/write **Long** .|
|[LockAspectRatio](shape-lockaspectratio-property-word.md)| **MsoTrue** if the specified shape retains its original proportions when you resize it. **MsoFalse** if you can change the height and width of the shape independently of one another when you resize it. Read/write **MsoTriState** .|
|[Name](shape-name-property-word.md)|Returns or sets the name of the specified object. Read/write  **String** .|
|[Nodes](shape-nodes-property-word.md)|Returns a  **[ShapeNodes](shapenodes-object-word.md)** collection that represents the geometric description of the specified shape.|
|[OLEFormat](shape-oleformat-property-word.md)|Returns an  **OLEFormat** object that represents the OLE characteristics (other than linking) for the specified shape, inline shape, or field. Read-only.|
|[Parent](shape-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **Shape** object.|
|[ParentGroup](shape-parentgroup-property-word.md)|Returns a  **Shape** object that represents the common parent shape of a child shape or a range of child shapes.|
|[PictureFormat](shape-pictureformat-property-word.md)|Returns a  **PictureFormat** object that contains picture formatting properties for the specified object. Read-only.|
|[Reflection](shape-reflection-property-word.md)|Returns a  **[ReflectionFormat](reflectionformat-object-word.md)** object that represents the reflection formatting for a shape. Read-only.|
|[RelativeHorizontalPosition](shape-relativehorizontalposition-property-word.md)|Specifies to the relative horizontal position of a shape. Read/write  **[WdRelativeHorizontalPosition](wdrelativehorizontalposition-enumeration-word.md)** .|
|[RelativeHorizontalSize](shape-relativehorizontalsize-property-word.md)|Returns or sets a  **[WdRelativeVerticalSize](wdrelativeverticalsize-enumeration-word.md)** constant that represents the object to which a range of shapes is relative. Read/write.|
|[RelativeVerticalPosition](shape-relativeverticalposition-property-word.md)|Specifies the relative vertical position of a shape. Read/write  **WdRelativeVerticalPosition** .|
|[RelativeVerticalSize](shape-relativeverticalsize-property-word.md)|Returns or sets a  **[WdRelativeVerticalSize](wdrelativeverticalsize-enumeration-word.md)** constant that represents the relative vertical size of a shape. Read/write.|
|[Rotation](shape-rotation-property-word.md)|Returns or sets the number of degrees the specified shape is rotated around the z-axis. A positive value indicates clockwise rotation; a negative value indicates counterclockwise rotation. Read/write  **Single** .|
|[Script](shape-script-property-word.md)|Returns a  **Script** object, which represents a block of script or code for an image on a Web page.|
|[Shadow](shape-shadow-property-word.md)|Returns a  **ShadowFormat** object that represents the shadow formatting for the specified shape.|
|[ShapeStyle](shape-shapestyle-property-word.md)|Returns or sets the shape style for the specified shape. Read/write  **[MsoShapeStyleIndex](msoshapestyleindex-enumeration-office.md)** .|
|[SmartArt](shape-smartart-property-word.md)|Returns a [SmartArt](smartart-object-office.md) object that provides a way to work with the SmartArt associated with the specified shape. Read-only.|
|[SoftEdge](shape-softedge-property-word.md)|Returns a  **[SoftEdgeFormat](softedgeformat-object-word.md)** object that represents the soft edge formatting for a shape. Read-only.|
|[TextEffect](shape-texteffect-property-word.md)|Returns a  **TextEffectFormat** object that contains text-effect formatting properties for the specified shape. Read-only.|
|[TextFrame](shape-textframe-property-word.md)|Returns a  **TextFrame** object that contains the text for the specified shape.|
|[TextFrame2](shape-textframe2-property-word.md)|Returns a  **TextFrame2** object that contains the text for the specified shape. Read-only.|
|[ThreeD](shape-threed-property-word.md)|Returns a  **ThreeDFormat** object that contains 3-D formatting properties for the specified shape. Read-only.|
|[Title](shape-title-property-word.md)|Returns or sets a  **String** that contains a title for the specified shape. Read/write.|
|[Top](shape-top-property-word.md)|Returns or sets the vertical position of the specified shape or shape range in points. Read/write  **Single** .|
|[TopRelative](shape-toprelative-property-word.md)|Returns or sets a  **Single** that represents the relative top position of a shape. Read/write.|
|[Type](shape-type-property-word.md)|Returns the type of inline shape. Read-only  **MsoShapeType** .|
|[VerticalFlip](shape-verticalflip-property-word.md)| **True** if the specified shape is flipped around the vertical axis. Read-only **MsoTriState** .|
|[Vertices](shape-vertices-property-word.md)|Returns the coordinates of the specified freeform drawing's vertices (and control points for B?zier curves) as a series of coordinate pairs. Read-only  **Variant** .|
|[Visible](shape-visible-property-word.md)| **True** if the specified object, or the formatting applied to it, is visible. Read/write **MsoTriState** .|
|[Width](shape-width-property-word.md)|Returns or sets the width, in points, of the specified shape. Read/write  **Long** .|
|[WidthRelative](shape-widthrelative-property-word.md)|Returns or sets a  **Single** that represents the relative width of a shape. Read/write.|
|[WrapFormat](shape-wrapformat-property-word.md)|Returns a  **WrapFormat** object that contains the properties for wrapping text around the specified shape. Read-only.|
|[ZOrderPosition](shape-zorderposition-property-word.md)|Returns a  **Long** that represents the position of the specified shape in the z-order. Read-only.|

