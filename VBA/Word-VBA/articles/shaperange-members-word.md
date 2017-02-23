---
title: ShapeRange Members (Word)
ms.prod: WORD
ms.assetid: eb882d13-d724-26e9-7e6d-2af55e42bba1
---


# ShapeRange Members (Word)
Represents a shape range, which is a set of shapes on a document. A shape range can contain as few as one shape or as many as all the shapes in the document. 

Represents a shape range, which is a set of shapes on a document. A shape range can contain as few as one shape or as many as all the shapes in the document. 


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Align](shaperange-align-method-word.md)|Aligns the shapes in the specified range of shapes.|
|[Apply](shaperange-apply-method-word.md)|Applies to the specified shape formatting that has been copied using the  **PickUp** method.|
|[CanvasCropBottom](shaperange-canvascropbottom-method-word.md)|Crops a percentage of the height of a drawing canvas from the bottom of the canvas.|
|[CanvasCropLeft](shaperange-canvascropleft-method-word.md)|Crops a percentage of the width of a drawing canvas from the left side of the canvas.|
|[CanvasCropRight](shaperange-canvascropright-method-word.md)|Crops a percentage of the width of a drawing canvas from the right side of the canvas.|
|[CanvasCropTop](shaperange-canvascroptop-method-word.md)|Crops a percentage of the height of a drawing canvas from the top of the canvas.|
|[ConvertToInlineShape](shaperange-converttoinlineshape-method-word.md)|Converts the specified shape in the drawing layer of a document to an inline shape in the text layer. You can convert only shapes that represent pictures, OLE objects, or ActiveX controls. .|
|[Delete](shaperange-delete-method-word.md)|Deletes the specified range of shapes.|
|[Distribute](shaperange-distribute-method-word.md)|Evenly distributes the shapes in the specified range of shapes. .|
|[Duplicate](shaperange-duplicate-method-word.md)|Creates a duplicate of the specified  **ShapeRange** object, adds the new range of shapes to the **Shapes** collection at a standard offset from the original shapes, and then returns a **Shape** object.|
|[Flip](shaperange-flip-method-word.md)|Flips a shape horizontally or vertically.|
|[Group](shaperange-group-method-word.md)|Groups the shapes in the specified range, and returns the grouped shapes as a single  **Shape** object.|
|[IncrementLeft](shaperange-incrementleft-method-word.md)|Moves the specified shape horizontally by the specified number of points.|
|[IncrementRotation](shaperange-incrementrotation-method-word.md)|Changes the rotation of the specified shape around the z-axis by the specified number of degrees. .|
|[IncrementTop](shaperange-incrementtop-method-word.md)|Moves the specified shape vertically by the specified number of points.|
|[Item](shaperange-item-method-word.md)|Returns an individual  **Shape** object in a collection.|
|[PickUp](shaperange-pickup-method-word.md)|Copies the formatting of the specified shape.|
|[ScaleHeight](shaperange-scaleheight-method-word.md)|Scales the height of a range of shapes by a specified factor.|
|[ScaleWidth](shaperange-scalewidth-method-word.md)|Scales the width of a shape by a specified factor.|
|[Select](shaperange-select-method-word.md)|Selects the specified range of shapes.|
|[SetShapesDefaultProperties](shaperange-setshapesdefaultproperties-method-word.md)|Applies the formatting of a default shape for a document to the specified range of shapes.|
|[Ungroup](shaperange-ungroup-method-word.md)|Ungroups any grouped shapes in the specified range of shapes, disassembles pictures and OLE objects within the specified shape or range of shapes, and returns the ungrouped shapes as a single  **ShapeRange** object.|
|[ZOrder](shaperange-zorder-method-word.md)|Moves the specified shape range in front of or behind other shapes in the collection (that is, changes the shape range's position in the z-order).|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Adjustments](shaperange-adjustments-property-word.md)|Returns an  **[Adjustments](adjustments-object-word.md)** object that contains adjustment values for all the adjustments in the specified **ShapeRange** object that represents an AutoShape or WordArt. Read-only.|
|[AlternativeText](shaperange-alternativetext-property-word.md)|Returns or sets the alternative text associated with a shape in a Web page. Read/write  **String** .|
|[Anchor](shaperange-anchor-property-word.md)|Returns a  **Range** object that represents the anchoring range for the specified shape range. Read-only.|
|[Application](shaperange-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[AutoShapeType](shaperange-autoshapetype-property-word.md)|Returns or sets the shape type for the specified  **ShapeRange** object, which must represent an AutoShape other than a line or freeform drawing. Read/write **MsoAutoShapeType** .|
|[BackgroundStyle](shaperange-backgroundstyle-property-word.md)|Sets or returns the background style of the shapes in the specified shape range. Read/write [MsoBackgroundStyleIndex](msobackgroundstyleindex-enumeration-office.md).|
|[Callout](shaperange-callout-property-word.md)|Returns a  **[CalloutFormat](calloutformat-object-word.md)** object that contains callout formatting properties for the specified shape. Read-only.|
|[CanvasItems](shaperange-canvasitems-property-word.md)|Returns a  **[CanvasShapes](canvasshapes-object-word.md)** object that represents a collection of shapes in a drawing canvas.|
|[Child](shaperange-child-property-word.md)| **True** if the shape is a child shape or if all shapes in a shape range are child shapes of the same parent. Read-only **MsoTriState** .|
|[Count](shaperange-count-property-word.md)|Returns a  **Long** that represents the number of shapes in the collection. Read-only.|
|[Creator](shaperange-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[Fill](shaperange-fill-property-word.md)|Returns a  **[FillFormat](fillformat-object-word.md)** object that contains fill formatting properties for the specified shape. Read-only.|
|[Glow](shaperange-glow-property-word.md)|Returns a  **[GlowFormat](glowformat-object-word.md)** object that represents the glow formatting for a range of shapes. Read-only.|
|[GroupItems](shaperange-groupitems-property-word.md)|Returns a  **[GroupShapes](groupshapes-object-word.md)** object that represents the individual shapes in the specified group. Read-only.|
|[Height](shaperange-height-property-word.md)|Returns or sets the height of the specified shape range. Read/write  **Single** .|
|[HeightRelative](shaperange-heightrelative-property-word.md)|Returns or sets a  **Single** that represents the percentage of the target shape to which the range of shapes is sized. Read/write.|
|[HorizontalFlip](shaperange-horizontalflip-property-word.md)|Indicates that a range of shapes has been flipped horizontally. Read-only  **MsoTriState** .|
|[Hyperlink](shaperange-hyperlink-property-word.md)|Returns a  **Hyperlink** object that represents the hyperlink associated with the specified **ShapeRange** object. Read-only.|
|[ID](shaperange-id-property-word.md)|Returns the identification type for the range of shapes. Read-only  **Long** .|
|[LayoutInCell](shaperange-layoutincell-property-word.md)|Returns a  **Long** that represents whether a shape in a table is displayed inside the table or outside the table. .|
|[Left](shaperange-left-property-word.md)|Returns or sets a  **Single** that represents the horizontal position, measured in points, of the specified range of shapes. Can also be any valid **[WdShapePosition](wdshapeposition-enumeration-word.md)** constant. Read/write.|
|[LeftRelative](shaperange-leftrelative-property-word.md)|Returns or sets a  **Single** that represents the relative left position of a range of shapes. Read/write.|
|[Line](shaperange-line-property-word.md)|Returns a  **LineFormat** object that contains line formatting properties for the specified range of shapes. Read-only.|
|[LockAnchor](shaperange-lockanchor-property-word.md)| **True** if the anchor for the specified **ShapeRange** object is locked to the anchoring range. Read/write **Long** .|
|[LockAspectRatio](shaperange-lockaspectratio-property-word.md)| **MsoTrue** if the specified shape retains its original proportions when you resize it. **MsoFalse** if you can change the height and width of the shape independently of one another when you resize it. Read/write **MsoTriState** .|
|[Name](shaperange-name-property-word.md)|Returns or sets the name of the specified object. Read/write  **String** .|
|[Nodes](shaperange-nodes-property-word.md)|Returns a  **ShapeNodes** collection that represents the geometric description of the specified shape.|
|[Parent](shaperange-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **ShapeRange** object.|
|[ParentGroup](shaperange-parentgroup-property-word.md)|Returns a  **Shape** object that represents the common parent shape of a range of shapes.|
|[PictureFormat](shaperange-pictureformat-property-word.md)|Returns a  **PictureFormat** object that contains picture formatting properties for the specified range of shapes. Read-only.|
|[Reflection](shaperange-reflection-property-word.md)|Returns a  **[ReflectionFormat](reflectionformat-object-word.md)** object that represents the reflection formatting for a range of shapes. Read-only.|
|[RelativeHorizontalPosition](shaperange-relativehorizontalposition-property-word.md)|Specifies the relative horizontal position of a range of shapes. Read/write  **[WdRelativeHorizontalPosition](wdrelativehorizontalposition-enumeration-word.md)** .|
|[RelativeHorizontalSize](shaperange-relativehorizontalsize-property-word.md)|Returns or sets a  **[WdRelativeHorizontalSize](wdrelativehorizontalsize-enumeration-word.md)** constant that represents the object to which a range of shapes is relative. Read/write.|
|[RelativeVerticalPosition](shaperange-relativeverticalposition-property-word.md)|Specifies the relative vertical position of a range of shapes. Read/write **[WdRelativeHorizontalPosition](wdrelativehorizontalposition-enumeration-word.md)** .|
|[RelativeVerticalSize](shaperange-relativeverticalsize-property-word.md)|Returns or sets a  **[WdRelativeVerticalSize](wdrelativeverticalsize-enumeration-word.md)** constant that represents the object to which a range of shapes is relative. Read/write.|
|[Rotation](shaperange-rotation-property-word.md)|Returns or sets the number of degrees the specified shape is rotated around the z-axis. Read/write  **Single** .|
|[Shadow](shaperange-shadow-property-word.md)|Returns a  **ShadowFormat** object that represents the shadow formatting for the specified shape.|
|[ShapeStyle](shaperange-shapestyle-property-word.md)|Returns or sets the shape style for the shapes in the specified shape range. Read/write [MsoShapeStyleIndex](msoshapestyleindex-enumeration-office.md).|
|[SoftEdge](shaperange-softedge-property-word.md)|Returns a  **[SoftEdgeFormat](softedgeformat-object-word.md)** object that represents the soft edge formatting for a range of shapes. Read-only.|
|[TextEffect](shaperange-texteffect-property-word.md)|Returns a  **TextEffectFormat** object that contains text-effect formatting properties for the specified shape. Read-only.|
|[TextFrame](shaperange-textframe-property-word.md)|Returns a  **TextFrame** object that contains the text for the specified range of shapes.|
|[TextFrame2](shaperange-textframe2-property-word.md)|Returns a  **TextFrame2** object that contains the text for the specified range of shapes. Read-only.|
|[ThreeD](shaperange-threed-property-word.md)|Returns a  **ThreeDFormat** object that contains 3-D formatting properties for the specified range of shapes. Read-only.|
|[Title](shaperange-title-property-word.md)|Returns or sets a  **String** that contains a title for the shapes in the specified shape range. Read/write.|
|[Top](shaperange-top-property-word.md)|Returns or sets the vertical position of the specified shape or shape range in points. Read/write  **Single** .|
|[TopRelative](shaperange-toprelative-property-word.md)|Returns or sets a  **Single** that represents the relative top position of a range of shapes. Read/write.|
|[Type](shaperange-type-property-word.md)|Returns the shape type. Read-only  **MsoShapeType** .|
|[VerticalFlip](shaperange-verticalflip-property-word.md)| **True** if the specified shape is flipped around the vertical axis. Read-only **MsoTriState** .|
|[Vertices](shaperange-vertices-property-word.md)|Returns the coordinates of the specified freeform drawing's vertices (and control points for B?zier curves) as a series of coordinate pairs. You can use the array returned by this property as an argument for the  **AddCurve** or **AddPolyLine** method. Read-only **Variant** .|
|[Visible](shaperange-visible-property-word.md)| **True** if the specified object, or the formatting applied to it, is visible. Read/write **MsoTriState** .|
|[Width](shaperange-width-property-word.md)|Returns or sets the width, in points, of the shapes within the range. Read/write  **Long** .|
|[WidthRelative](shaperange-widthrelative-property-word.md)|Returns or sets a  **Single** that represents the relative width of a range of shapes. Read/write.|
|[WrapFormat](shaperange-wrapformat-property-word.md)|Returns a  **WrapFormat** object that contains the properties for wrapping text around the specified range of shapes. Read-only.|
|[ZOrderPosition](shaperange-zorderposition-property-word.md)|Returns a  **Long** that represents the position of the specified shape in the z-order. Read-only.|

