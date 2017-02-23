---
title: Shape Members (PowerPoint)
ms.prod: POWERPOINT
ms.assetid: e371c375-c16a-33ef-32b7-6dcb99d3d128
---


# Shape Members (PowerPoint)
Represents an object in the drawing layer, such as an AutoShape, freeform, OLE object, or picture.

Represents an object in the drawing layer, such as an AutoShape, freeform, OLE object, or picture.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Apply](shape-apply-method-powerpoint.md)|Applies to the specified shape formatting that's been copied by using the  **PickUp** method.|
|[ApplyAnimation](shape-applyanimation-method-powerpoint.md)|Applies the last picked up animation to the  **Shape** object.|
|[ConvertTextToSmartArt](shape-converttexttosmartart-method-powerpoint.md)|Converts text in a  **Shape** object to a SmartArt diagram.|
|[Copy](shape-copy-method-powerpoint.md)|Copies the specified object to the Clipboard.|
|[Cut](shape-cut-method-powerpoint.md)|Deletes the specified object and places it on the Clipboard.|
|[Delete](shape-delete-method-powerpoint.md)|Deletes the specified  **Shape** object.|
|[Duplicate](shape-duplicate-method-powerpoint.md)|Creates a duplicate of the specified  **Shape** object, adds the new shape to the **Shapes** collection, and then returns a new **ShapeRange** object. The duplicated objects are placed at the end of the **Shapes** collection.|
|[Flip](shape-flip-method-powerpoint.md)|Flips the specified shape around its horizontal or vertical axis.|
|[IncrementLeft](shape-incrementleft-method-powerpoint.md)|Moves the specified shape horizontally by the specified number of points.|
|[IncrementRotation](shape-incrementrotation-method-powerpoint.md)|Changes the rotation of the specified shape around the z-axis by the specified number of degrees. Use the  **Rotation** property to set the absolute rotation of the shape.|
|[IncrementTop](shape-incrementtop-method-powerpoint.md)|Moves the specified shape vertically by the specified number of points.|
|[PickUp](shape-pickup-method-powerpoint.md)|Copies the formatting of the specified shape. Use the  **Apply** method to apply the copied formatting to another shape.|
|[PickupAnimation](shape-pickupanimation-method-powerpoint.md)|Picks up all animation from the  **Shape** object.|
|[RerouteConnections](shape-rerouteconnections-method-powerpoint.md)|Reroutes connectors so that they take the shortest possible path between the shapes they connect. To do this, the  **RerouteConnections** method may detach the ends of a connector and reattach them to different connecting sites on the connected shapes.|
|[ScaleHeight](shape-scaleheight-method-powerpoint.md)|Scales the height of the shape by a specified factor.|
|[ScaleWidth](shape-scalewidth-method-powerpoint.md)|Scales the width of the shape by a specified factor. For pictures and OLE objects, you can indicate whether you want to scale the shape relative to the original size or relative to the current size. Shapes other than pictures and OLE objects are always scaled relative to their current width.|
|[Select](shape-select-method-powerpoint.md)|Selects the specified object.|
|[SetShapesDefaultProperties](shape-setshapesdefaultproperties-method-powerpoint.md)|Applies the formatting for the specified shape to the default shape. Shapes created after this method has been used will have this formatting applied to them by default.|
|[Ungroup](shape-ungroup-method-powerpoint.md)|Ungroups any grouped shapes in the specified shape or range of shapes. Disassembles pictures and OLE objects within the specified shape or range of shapes. Returns the ungrouped shapes as a single  **[ShapeRange](shaperange-object-powerpoint.md)** object.|
|[UpgradeMedia](shape-upgrademedia-method-powerpoint.md)|Converts a legacy media object to an updated media object.|
|[ZOrder](shape-zorder-method-powerpoint.md)|Moves the specified shape in front of or behind other shapes in the collection (that is, changes the shape's position in the z-order).|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[ActionSettings](shape-actionsettings-property-powerpoint.md)|Returns an  **[ActionSettings](actionsettings-object-powerpoint.md)** object that contains information about what action occurs when the user clicks or moves the mouse over the specified shape or text range during a slide show. Read-only.|
|[Adjustments](shape-adjustments-property-powerpoint.md)|Returns an  **[Adjustments](adjustments-object-powerpoint.md)** object that contains adjustment values for all the adjustments in the specified shape. Applies to any **Shape** object that represents an AutoShape, WordArt, or a connector. Read-only.|
|[AlternativeText](shape-alternativetext-property-powerpoint.md)|Returns or sets the alternative text associated with a shape in a Web presentation. Read/write.|
|[AnimationSettings](shape-animationsettings-property-powerpoint.md)|Returns an  **[AnimationSettings](animationsettings-object-powerpoint.md)** object that represents all the special effects you can apply to the animation of the specified shape. Read-only.|
|[Application](shape-application-property-powerpoint.md)|Returns an  **[Application](application-object-powerpoint.md)** object that represents the creator of the specified object.|
|[AutoShapeType](shape-autoshapetype-property-powerpoint.md)|Returns or sets the shape type for the specified  **Shape** object, which must represent an AutoShape other than a line, freeform drawing, or connector. Read/write.|
|[BackgroundStyle](shape-backgroundstyle-property-powerpoint.md)|Sets or returns the background style of the specified object. Read/write.|
|[BlackWhiteMode](shape-blackwhitemode-property-powerpoint.md)|Returns or sets a value that indicates how the specified shape appears when the presentation is viewed in black-and-white mode. Read/write.|
|[Callout](shape-callout-property-powerpoint.md)|Returns a  **[CalloutFormat](calloutformat-object-powerpoint.md)** object that contains callout formatting properties for the specified shape. Applies to **Shape** or **ShapeRange** objects that represent line callouts. Read-only.|
|[Chart](shape-chart-property-powerpoint.md)|Returns a  **Chart** object of the current **Shape** object. Read-only.|
|[Child](shape-child-property-powerpoint.md)|**MsoTrue** if the shape is a child shape or if all shapes in a shape range are child shapes of the same parent. Read-only.|
|[ConnectionSiteCount](shape-connectionsitecount-property-powerpoint.md)|Returns the number of connection sites on the specified shape. Read-only.|
|[Connector](shape-connector-property-powerpoint.md)|Determines whether the specified shape is a connector. Read-only.|
|[ConnectorFormat](shape-connectorformat-property-powerpoint.md)|Returns a  **[ConnectorFormat](connectorformat-object-powerpoint.md)** object that contains connector formatting properties. Applies to **Shape** or **ShapeRange** objects that represent connectors. Read-only.|
|[Creator](shape-creator-property-powerpoint.md)|Returns a  **Long** that represents the four-character creator code for the application in which the specified object was created. For example, if the object was created in Microsoft PowerPoint, this property returns the hexadecimal number 50575054. Read-only.|
|[CustomerData](shape-customerdata-property-powerpoint.md)|Returns a  **[CustomerData](customerdata-object-powerpoint.md)** object. Read-only.|
|[Fill](shape-fill-property-powerpoint.md)|Returns a  **[FillFormat](fillformat-object-powerpoint.md)** object that contains fill formatting properties for the specified shape. Read-only.|
|[Glow](shape-glow-property-powerpoint.md)|Returns the glow format for the specified shape. Read-only.|
|[GroupItems](shape-groupitems-property-powerpoint.md)|Returns a  **[GroupShapes](groupshapes-object-powerpoint.md)** object that represents the individual shapes in the specified group. Use the **Item** method of the **GroupShapes** object to return a single shape from the group. Read-only.|
|[HasChart](shape-haschart-property-powerpoint.md)|Returns whether the shape represented by the specified object contains a chart. Read-only.|
|[HasInkXML](shape-hasinkxml-property-powerpoint.md)|Returns an [MsoTriState](msotristate-enumeration-office.md) enumeration value that indicates whether the specified shape contains ink XML that can be retrieved via the[Shape.InkXML](shape-inkxml-property-powerpoint.md) property. Read-only.|
|[HasSmartArt](shape-hassmartart-property-powerpoint.md)|Returns  **True** if the current **Shape** object contains a SmartArt diagram. Read-only.|
|[HasTable](shape-hastable-property-powerpoint.md)|Returns whether the specified shape is a table. Read-only.|
|[HasTextFrame](shape-hastextframe-property-powerpoint.md)|Returns whether the specified shape has a text frame. Read-only.|
|[Height](shape-height-property-powerpoint.md)|Returns or sets the height of the specified object, in points. Read/write.|
|[HorizontalFlip](shape-horizontalflip-property-powerpoint.md)|Returns whether the specified shape is flipped around the horizontal axis. Read-only.|
|[Id](shape-id-property-powerpoint.md)|Returns a  **Long** that identifies the shape or range of shapes. Read-only.|
|[InkXML](shape-inkxml-property-powerpoint.md)|Returns a  **String** that contains the InkActionML associated with the specified shape. Read-only.|
|[IsNarration](shape-isnarration-property-powerpoint.md)|Specifies whether the specified shape range contains a narration. Read/write. |
|[Left](shape-left-property-powerpoint.md)|Returns or sets a  **Single** that represents the distance in points from the left edge of the shape's bounding box to the left edge of the slide. Read/write.|
|[Line](shape-line-property-powerpoint.md)|Returns a  **[LineFormat](lineformat-object-powerpoint.md)** object that contains line formatting properties for the specified shape. (For a line, the **LineFormat** object represents the line itself; for a shape with a border, the **LineFormat** object represents the border.) Read-only.|
|[LinkFormat](shape-linkformat-property-powerpoint.md)|Returns a  **[LinkFormat](linkformat-object-powerpoint.md)** object that contains the properties that are unique to linked OLE objects. Read-only.|
|[LockAspectRatio](shape-lockaspectratio-property-powerpoint.md)|Determines whether the specified shape retains its original proportions when you resize it. Read/write.|
|[MediaFormat](shape-mediaformat-property-powerpoint.md)|Allows access to the new audio or video object. Read-only.|
|[MediaType](shape-mediatype-property-powerpoint.md)|Returns the OLE media type. Read-only.|
|[Name](shape-name-property-powerpoint.md)|When a shape is created, Microsoft PowerPoint automatically assigns it a name in the form  _ShapeType Number_, where _ShapeType_ identifies the type of shape or AutoShape, and _Number_ is an integer that's unique within the collection of shapes on the slide. For example, the automatically generated names of the shapes on a slide could be Placeholder 1, Oval 2, and Rectangle 3. To avoid conflict with automatically assigned names, don't use the form _ShapeType Number_ for user-defined names, where _ShapeType_ is a value that is used for automatically generated names, and _Number_ is any positive integer. A shape range must contain exactly one shape. Read/write.|
|[Nodes](shape-nodes-property-powerpoint.md)|Returns a  **[ShapeNodes](shapenodes-object-powerpoint.md)** collection that represents the geometric description of the specified shape. Applies to **Shape** objects that represent freeform drawings.|
|[OLEFormat](shape-oleformat-property-powerpoint.md)|Returns an  **[OLEFormat](oleformat-object-powerpoint.md)** object that contains OLE formatting properties for the specified shape. Applies to **Shape** or **ShapeRange** objects that represent OLE objects. Read-only.|
|[Parent](shape-parent-property-powerpoint.md)|Returns the parent object for the specified object.|
|[ParentGroup](shape-parentgroup-property-powerpoint.md)|Returns a  **Shape** object that represents the common parent shape of a child shape or a range of child shapes.|
|[PictureFormat](shape-pictureformat-property-powerpoint.md)|Returns a  **[PictureFormat](pictureformat-object-powerpoint.md)** object that contains picture formatting properties for the specified shape. Read-only.|
|[PlaceholderFormat](shape-placeholderformat-property-powerpoint.md)|Returns a  **[PlaceholderFormat](placeholderformat-object-powerpoint.md)** object that contains the properties that are unique to placeholders. Read-only.|
|[Reflection](shape-reflection-property-powerpoint.md)|Returns the reflection format for the specified shape. Read-only.|
|[Rotation](shape-rotation-property-powerpoint.md)|Returns or sets the number of degrees the specified shape is rotated around the z-axis. Read/write.|
|[Shadow](shape-shadow-property-powerpoint.md)|Returns a  **[ShadowFormat](shadowformat-object-powerpoint.md)** object that contains shadow formatting properties for the specified shape. Read-only.|
|[ShapeStyle](shape-shapestyle-property-powerpoint.md)|Sets or returns the shape style index for the specified object. Read/write.|
|[SmartArt](shape-smartart-property-powerpoint.md)|Returns a Microsoft Office [SmartArt](smartart-object-office.md) object that represents the SmartArt diagram of the **Shape** object. Read-only.|
|[SoftEdge](shape-softedge-property-powerpoint.md)|Returns the soft edge format for the specified shape. Read-only.|
|[Table](shape-table-property-powerpoint.md)|Returns a  **[Table](table-object-powerpoint.md)** object that represents a table in a shape or in a shape range. Read-only.|
|[Tags](shape-tags-property-powerpoint.md)|Returns a  **[Tags](tags-object-powerpoint.md)** object that represents the tags for the specified object. Read-only.|
|[TextEffect](shape-texteffect-property-powerpoint.md)|Returns a  **[TextEffectFormat](texteffectformat-object-powerpoint.md)** object that contains text-effect formatting properties for the specified shape. Read-only.|
|[TextFrame](shape-textframe-property-powerpoint.md)|Returns a  **[TextFrame](textframe-object-powerpoint.md)** object that contains the alignment and anchoring properties for the specified shape or master text style.|
|[TextFrame2](shape-textframe2-property-powerpoint.md)|Returns the  **[TextFrame2](textframe2-object-powerpoint.md)** object associated with the specified **[Shape](shape-object-powerpoint.md)** object that contains the alignment and anchoring properties for the specified shape. Read-only.|
|[ThreeD](shape-threed-property-powerpoint.md)|Returns a  **[ThreeDFormat](threedformat-object-powerpoint.md)** object that contains 3-D - effect formatting properties for the specified shape. Read-only.|
|[Title](shape-title-property-powerpoint.md)|Returns a  **[Shape](shape-object-powerpoint.md)** object that represents the slide title. Read-only.|
|[Top](shape-top-property-powerpoint.md)|Returns or sets a  **Single** that represents the distance from the top edge of the shape's bounding box to the top edge of the document. Read/write.|
|[Type](shape-type-property-powerpoint.md)|Represents the type of shape or shapes in a range of shapes. Read-only.|
|[VerticalFlip](shape-verticalflip-property-powerpoint.md)|Determines whether the specified shape is flipped around the vertical axis. Read-only.|
|[Vertices](shape-vertices-property-powerpoint.md)|Returns the coordinates of the specified freeform drawing's vertices (and control points for B?zier curves) as a series of coordinate pairs. Read-only.|
|[Visible](shape-visible-property-powerpoint.md)|Returns or sets the visibility of the specified object or the formatting applied to the specified object. Read/write.|
|[Width](shape-width-property-powerpoint.md)|Returns or sets the width of the specified object, in points. Read/write.|
|[ZOrderPosition](shape-zorderposition-property-powerpoint.md)|Returns the position of the specified shape in the z-order. Read-only.|

