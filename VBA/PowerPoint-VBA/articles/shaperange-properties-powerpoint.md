---
title: ShapeRange Properties (PowerPoint)
ms.prod: POWERPOINT
ms.assetid: 8fc24857-8038-455c-ae7f-18508738e58b
---


# ShapeRange Properties (PowerPoint)

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[ActionSettings](shaperange-actionsettings-property-powerpoint.md)|Returns an  **[ActionSettings](actionsettings-object-powerpoint.md)** object that contains information about what action occurs when the user clicks or moves the mouse over the specified shape or text range during a slide show. Read-only.|
|[Adjustments](shaperange-adjustments-property-powerpoint.md)|Returns an  **[Adjustments](adjustments-object-powerpoint.md)** object that contains adjustment values for all the adjustments in the specified shape. Applies to any **ShapeRange** object that represents an AutoShape, WordArt, or a connector. Read-only.|
|[AlternativeText](shaperange-alternativetext-property-powerpoint.md)|Returns or sets the alternative text associated with a shape in a Web presentation. Read/write.|
|[AnimationSettings](shaperange-animationsettings-property-powerpoint.md)|Returns an  **[AnimationSettings](animationsettings-object-powerpoint.md)** object that represents all the special effects you can apply to the animation of the specified shape. Read-only.|
|[Application](shaperange-application-property-powerpoint.md)|Returns an  **[Application](application-object-powerpoint.md)** object that represents the creator of the specified object.|
|[AutoShapeType](shaperange-autoshapetype-property-powerpoint.md)|Returns or sets the shape type for the specified  **ShapeRange** object, which must represent an AutoShape other than a line, freeform drawing, or connector. Read/write.|
|[BackgroundStyle](shaperange-backgroundstyle-property-powerpoint.md)|Sets or returns the background style of the specified object. Read/write.|
|[BlackWhiteMode](shaperange-blackwhitemode-property-powerpoint.md)|Returns or sets a value that indicates how the specified shape appears when the presentation is viewed in black-and-white mode. Read/write.|
|[Callout](shaperange-callout-property-powerpoint.md)|Returns a  **[CalloutFormat](calloutformat-object-powerpoint.md)** object that contains callout formatting properties for the specified shape. Applies to **Shape** or **ShapeRange** objects that represent line callouts. Read-only.|
|[Chart](shaperange-chart-property-powerpoint.md)|Returns the  **Chart** object of the current **ShapeRange** object. Read-only.|
|[Child](shaperange-child-property-powerpoint.md)|**MsoTrue** if the shape is a child shape or if all shapes in a shape range are child shapes of the same parent. Read-only.|
|[ConnectionSiteCount](shaperange-connectionsitecount-property-powerpoint.md)|Returns the number of connection sites on the specified shape. Read-only.|
|[Connector](shaperange-connector-property-powerpoint.md)|Determines whether the specified shape is a connector. Read-only.|
|[ConnectorFormat](shaperange-connectorformat-property-powerpoint.md)|Returns a  **[ConnectorFormat](connectorformat-object-powerpoint.md)** object that contains connector formatting properties. Applies to **Shape** or **ShapeRange** objects that represent connectors. Read-only.|
|[Count](shaperange-count-property-powerpoint.md)|Returns the number of objects in the specified collection. Read-only.|
|[Creator](shaperange-creator-property-powerpoint.md)|Returns a  **Long** that represents the four-character creator code for the application in which the specified object was created. For example, if the object was created in Microsoft PowerPoint, this property returns the hexadecimal number 50575054. Read-only.|
|[CustomerData](shaperange-customerdata-property-powerpoint.md)|Returns a  **[CustomerData](customerdata-object-powerpoint.md)** object.|
|[Fill](shaperange-fill-property-powerpoint.md)|Returns a  **[FillFormat](fillformat-object-powerpoint.md)** object that contains fill formatting properties for the specified shape. Read-only.|
|[Glow](shaperange-glow-property-powerpoint.md)|Returns the glow format for the specified range of shapes. Read-only.|
|[GroupItems](shaperange-groupitems-property-powerpoint.md)|Returns a  **[GroupShapes](groupshapes-object-powerpoint.md)** object that represents the individual shapes in the specified group. Use the **Item** method of the **GroupShapes** object to return a single shape from the group. Read-only.|
|[HasChart](shaperange-haschart-property-powerpoint.md)|Returns whether the shape range represented by the specified object contains a chart. Read-only.|
|[HasInkXML](shaperange-hasinkxml-property-powerpoint.md)|Returns an [MsoTriState](msotristate-enumeration-office.md) enumeration value that indicates whether the specified shape range contains ink XML that can be retrieved via the[ShapeRange.InkXML](shaperange-inkxml-property-powerpoint.md) property. Read-only.|
|[HasSmartArt](shaperange-hassmartart-property-powerpoint.md)|Returns  **True** if the current **ShapeRange** object has a SmartArt diagram. Read-only.|
|[HasTable](shaperange-hastable-property-powerpoint.md)|Returns whether the specified shape is a table. Read-only.|
|[HasTextFrame](shaperange-hastextframe-property-powerpoint.md)|Returns whether the specified shape has a text frame. Read-only.|
|[Height](shaperange-height-property-powerpoint.md)|Returns or sets the height of the specified object, in points. Read/write.|
|[HorizontalFlip](shaperange-horizontalflip-property-powerpoint.md)|Returns whether the specified shape is flipped around the horizontal axis. Read-only.|
|[Id](shaperange-id-property-powerpoint.md)|Returns a  **Long** that identifies the shape or range of shapes. Read-only.|
|[InkXML](shaperange-inkxml-property-powerpoint.md)|Returns a  **String** that contains the InkActionML associated with the specified shape range. Read-only.|
|[IsNarration](shaperange-isnarration-property-powerpoint.md)|Specifies whether the specified shape range contains a narration. Read/write. |
|[Left](shaperange-left-property-powerpoint.md)|Returns or sets a  **Single** that represents the distance in points from the left edge of the leftmost shape in the shape range to the left edge of the slide. Read/write.|
|[Line](shaperange-line-property-powerpoint.md)|Returns a  **[LineFormat](lineformat-object-powerpoint.md)** object that contains line formatting properties for the specified shape. (For a line, the **LineFormat** object represents the line itself; for a shape with a border, the **LineFormat** object represents the border.) Read-only.|
|[LinkFormat](shaperange-linkformat-property-powerpoint.md)|Returns a  **[LinkFormat](linkformat-object-powerpoint.md)** object that contains the properties that are unique to linked OLE objects. Read-only.|
|[LockAspectRatio](shaperange-lockaspectratio-property-powerpoint.md)|Determines whether the specified shape retains its original proportions when you resize it. Read/write.|
|[MediaFormat](shaperange-mediaformat-property-powerpoint.md)|Returns the current  **MediaFormat** object. Read-only.|
|[MediaType](shaperange-mediatype-property-powerpoint.md)|Returns the OLE media type. Read-only.|
|[Name](shaperange-name-property-powerpoint.md)|When a shape is created, Microsoft PowerPoint automatically assigns it a name in the form  _ShapeType Number_, where _ShapeType_ identifies the type of shape or AutoShape, and _Number_ is an integer that's unique within the collection of shapes on the slide. For example, the automatically generated names of the shapes on a slide could be Placeholder 1, Oval 2, and Rectangle 3. To avoid conflict with automatically assigned names, don't use the form _ShapeType Number_ for user-defined names, where _ShapeType_ is a value that is used for automatically generated names, and _Number_ is any positive integer. A shape range must contain exactly one shape. Read/write.|
|[Nodes](shaperange-nodes-property-powerpoint.md)|Returns a  **[ShapeNodes](shapenodes-object-powerpoint.md)** collection that represents the geometric description of the specified shape. Applies to **ShapeRange** objects that represent freeform drawings.|
|[OLEFormat](shaperange-oleformat-property-powerpoint.md)|Returns an  **[OLEFormat](oleformat-object-powerpoint.md)** object that contains OLE formatting properties for the specified shape. Applies to **Shape** or **ShapeRange** objects that represent OLE objects. Read-only.|
|[Parent](shaperange-parent-property-powerpoint.md)|Returns the parent object for the specified object.|
|[ParentGroup](shaperange-parentgroup-property-powerpoint.md)|Returns a  **Shape** object that represents the common parent shape of a child shape or a range of child shapes.|
|[PictureFormat](shaperange-pictureformat-property-powerpoint.md)|Returns a  **[PictureFormat](pictureformat-object-powerpoint.md)** object that contains picture formatting properties for the specified shape. Read-only.|
|[PlaceholderFormat](shaperange-placeholderformat-property-powerpoint.md)|Returns a  **[PlaceholderFormat](placeholderformat-object-powerpoint.md)** object that contains the properties that are unique to placeholders. Read-only.|
|[Reflection](shaperange-reflection-property-powerpoint.md)|Returns the reflection format for the specified range of shapes. Read-only.|
|[Rotation](shaperange-rotation-property-powerpoint.md)|Returns or sets the number of degrees the specified shape is rotated around the z-axis. Read/write.|
|[Shadow](shaperange-shadow-property-powerpoint.md)|Returns a  **[ShadowFormat](shadowformat-object-powerpoint.md)** object that contains shadow formatting properties for the specified shapes. Read-only.|
|[ShapeStyle](shaperange-shapestyle-property-powerpoint.md)|Sets or returns the shape style index for the specified object.|
|[SmartArt](shaperange-smartart-property-powerpoint.md)|Returns the SmartArt diagram of the  **ShapeRange** object. Read-only.|
|[SoftEdge](shaperange-softedge-property-powerpoint.md)|Returns the soft edge format for the specified range of shapes. Read-only.|
|[Table](shaperange-table-property-powerpoint.md)|Returns a  **[Table](table-object-powerpoint.md)** object that represents a table in a shape or in a shape range. Read-only.|
|[Tags](shaperange-tags-property-powerpoint.md)|Returns a  **[Tags](tags-object-powerpoint.md)** object that represents the tags for the specified object. Read-only.|
|[TextEffect](shaperange-texteffect-property-powerpoint.md)|Returns a  **[TextEffectFormat](texteffectformat-object-powerpoint.md)** object that contains text-effect formatting properties for the specified shape. Read-only.|
|[TextFrame](shaperange-textframe-property-powerpoint.md)|Returns a  **[TextFrame](textframe-object-powerpoint.md)** object that contains the alignment and anchoring properties for the specified shape or master text style. Read-only.|
|[TextFrame2](shaperange-textframe2-property-powerpoint.md)|Returns the  **[TextFrame2](textframe2-object-powerpoint.md)** object associated with the specified **[ShapeRange](shaperange-object-powerpoint.md)** object that contains the alignment and anchoring properties for the specified shape range. Read-only.|
|[ThreeD](shaperange-threed-property-powerpoint.md)|Returns a  **[ThreeDFormat](threedformat-object-powerpoint.md)** object that contains 3-D - effect formatting properties for the specified shape. Read-only.|
|[Title](shaperange-title-property-powerpoint.md)|Returns a  **[Shape](shape-object-powerpoint.md)** object that represents the slide title. Read-only.|
|[Top](shaperange-top-property-powerpoint.md)|Returns or sets a  **Single** that represents the distance from the top edge of the topmost shape in the shape range to the top edge of the document. Read/write.|
|[Type](shaperange-type-property-powerpoint.md)|Represents the type of shape or shapes in a range of shapes. Read-only.|
|[VerticalFlip](shaperange-verticalflip-property-powerpoint.md)|Determines whether the specified shape is flipped around the vertical axis. Read-only.|
|[Vertices](shaperange-vertices-property-powerpoint.md)|Returns the coordinates of the specified freeform drawing's vertices (and control points for B?zier curves) as a series of coordinate pairs. Read-only.|
|[Visible](shaperange-visible-property-powerpoint.md)|Returns or sets the visibility of the specified object or the formatting applied to the specified object. Read/write.|
|[Width](shaperange-width-property-powerpoint.md)|Returns or sets the width of the specified object, in points. Read/write.|
|[ZOrderPosition](shaperange-zorderposition-property-powerpoint.md)|Returns the position of the specified shape in the z-order. Read-only.|

