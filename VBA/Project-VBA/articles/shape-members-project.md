---
title: Shape Members (Project)
ms.prod: PROJECTSERVER
ms.assetid: 4e1f1f6b-ecb4-42f7-a8da-8b5c417f6891
---


# Shape Members (Project)
Represents an object in a Project report, such as a chart, report table, text box, freeform drawing, or picture.

Represents an object in a Project report, such as a chart, report table, text box, freeform drawing, or picture.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Apply](shape-apply-method-project.md)|Applies formatting to a shape, where the formatting information has been copied by using the  **[PickUp](shape-pickup-method-project.md)** method.|
|[Copy](shape-copy-method-project.md)|Copies the shape to the Clipboard.|
|[Cut](shape-cut-method-project.md)|Cuts the shape to the Clipboard.|
|[Delete](shape-delete-method-project.md)|Deletes the shape.|
|[Duplicate](shape-duplicate-method-project.md)|Duplicates a shape and returns a reference to the copy.|
|[Flip](shape-flip-method-project.md)|Flips the shape around its horizontal or vertical axis.|
|[IncrementLeft](shape-incrementleft-method-project.md)|Moves the shape horizontally by the specified number of points.|
|[IncrementRotation](shape-incrementrotation-method-project.md)|Rotates the shape around the z-axis by the specified number of degrees.|
|[IncrementTop](shape-incrementtop-method-project.md)|Moves the shape vertically by the specified number of points.|
|[PickUp](shape-pickup-method-project.md)|Copies the formatting of a shape.|
|[RerouteConnections](shape-rerouteconnections-method-project.md)|The  **RerouteConnections** method is not implemented in Project.|
|[ScaleHeight](shape-scaleheight-method-project.md)|Scales the height of the shape by a specified factor.|
|[ScaleWidth](shape-scalewidth-method-project.md)|Scales the width of the shape by a specified factor.|
|[Select](shape-select-method-project.md)|Selects the shape.|
|[SetShapesDefaultProperties](shape-setshapesdefaultproperties-method-project.md)|Applies the formatting of a default shape to the shape.|
|[Ungroup](shape-ungroup-method-project.md)|The  **Ungroup** method is not implemented in Project.|
|[ZOrder](shape-zorder-method-project.md)|Moves the shape in front of or behind other shapes (that is, changes the position in the z-order).|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Adjustments](shape-adjustments-property-project.md)|Gets an  **[Adjustments](http://msdn.microsoft.com/en-us/library/office/ff838852%28v=office.15%29)** object that contains adjustment values for all the adjustments in the shape. Applies to any **Shape** object that represents an AutoShape, WordArt, or a connector. Read-only **Adjustments**.|
|[AlternativeText](shape-alternativetext-property-project.md)|Gets or sets the descriptive (alternative) text string for a  **Shape** object when the object is saved to a web page. Read/write **String**.|
|[Application](shape-application-property-project.md)|Gets the  **[Application Object (Project)](application-object-project.md)** object. Read-only **Application**.|
|[AutoShapeType](shape-autoshapetype-property-project.md)|Gets or sets the shape type for the  **Shape** object, which must represent an AutoShape other than a line, freeform drawing, or connector. Read/write **MsoAutoShapeType**.|
|[BackgroundStyle](shape-backgroundstyle-property-project.md)|Gets or sets the background style. Read/write  **MsoBackgroundStyleIndex**.|
|[BlackWhiteMode](shape-blackwhitemode-property-project.md)|Gets or sets a value that indicates how the shape appears when it is viewed in black-and-white mode. Read/write  **MsoBlackWhiteMode**.|
|[Callout](shape-callout-property-project.md)|Gets callout formatting properties for the shape, when the  **Shape** object represents a callout. Read-only **CalloutFormat**.|
|[Chart](shape-chart-property-project.md)|Gets a  **Chart** object that represents the chart contained in the shape. Read-only **Chart**.|
|[Child](shape-child-property-project.md)|Gets a value that indicates whether the shape is a child shape. Read-only  **[MsoTriState](http://msdn.microsoft.com/en-us/library/office/ff860737%28v=office.15%29)**.|
|[ConnectionSiteCount](shape-connectionsitecount-property-project.md)|Gets the number of connection sites on the shape. Read-only  **Long**.|
|[Connector](shape-connector-property-project.md)|Gets a value that indicates whether the shape is a connector. Read-only  **[MsoTriState](http://msdn.microsoft.com/en-us/library/office/ff860737%28v=office.15%29)**|
|[ConnectorFormat](shape-connectorformat-property-project.md)|Gets a  **ConnectorFormat** object that contains connector formatting properties. Applies to a **Shape** that represents a connector. Read-only **[ConnectorFormat](http://msdn.microsoft.com/en-us/library/office/ff820940%28v=office.15%29)**.|
|[Fill](shape-fill-property-project.md)|Gets a  **FillFormat** object for the shape, if the shape contains fill formatting properties. Read-only **[FillFormat](http://msdn.microsoft.com/en-us/library/office/ff838198%28v=office.15%29)**.|
|[Glow](shape-glow-property-project.md)|Gets a  **GlowFormat** object for the shape, if the shape contains glow formatting properties. Read-only **[GlowFormat](http://msdn.microsoft.com/en-us/library/office/ff864010%28v=office.15%29)**.|
|[GroupItems](shape-groupitems-property-project.md)|Gets a  **GroupShapes** object that represents the individual shapes in a group, if the **Shape** object represents a group of shapes. Read-only **[GroupShapes](http://msdn.microsoft.com/en-us/library/office/ff195331%28v=office.15%29)**.|
|[HasChart](shape-haschart-property-project.md)|Gets a value that indicates whether the shape contains a chart. Read-only  **[MsoTriState](http://msdn.microsoft.com/en-us/library/office/ff860737%28v=office.15%29)**.|
|[HasTable](shape-hastable-property-project.md)|Gets a value that indicates whether the shape contains a table. Read-only  **[MsoTriState](http://msdn.microsoft.com/en-us/library/office/ff860737%28v=office.15%29)**.|
|[Height](shape-height-property-project.md)|Gets or sets the height of the shape, in points. Read-write  **Single**.|
|[HorizontalFlip](shape-horizontalflip-property-project.md)|Gets a value that indicates whether the shape is flipped around the horizontal axis. Read-only  **[MsoTriState](http://msdn.microsoft.com/en-us/library/office/ff860737%28v=office.15%29)**.|
|[ID](shape-id-property-project.md)|Gets the identification type for the shape. Read-only  **Long**.|
|[Left](shape-left-property-project.md)|Gets or sets the horizontal distance, in points, from the left side of the report to the left edge of the shape. Read-write  **Single**.|
|[Line](shape-line-property-project.md)|Gets the line formatting properties for the shape. Read-only  **[LineFormat](http://msdn.microsoft.com/en-us/library/office/ff194214%28v=office.15%29)**.|
|[LockAspectRatio](shape-lockaspectratio-property-project.md)|Gets or sets a value that indicates whether the shape retains its original proportions when you resize it; that is, whether the aspect ratio of the shape is locked. Read-write  **[MsoTriState](http://msdn.microsoft.com/en-us/library/office/ff860737%28v=office.15%29)**|
|[Name](shape-name-property-project.md)|Gets or sets the name of the shape. Read-write  **String**.|
|[Nodes](shape-nodes-property-project.md)|Gets the geometric description of nodes or control points in the shape. Read-only  **[ShapeNodes](http://msdn.microsoft.com/en-us/library/office/ff822109%28v=office.15%29)**.|
|[Parent](shape-parent-property-project.md)|Gets the parent object of the shape. Read-only  **Object**.|
|[ParentGroup](shape-parentgroup-property-project.md)|Gets the common parent shape of a child shape or a range of child shapes. Read-only  **Shape**.|
|[Reflection](shape-reflection-property-project.md)|Gets the reflection formatting for the shape. Read-only  **[ReflectionFormat](http://msdn.microsoft.com/en-us/library/office/ff863140%28v=office.15%29)**.|
|[Rotation](shape-rotation-property-project.md)|Gets or sets the rotation of the shape, in degrees. Read/write  **Single**.|
|[Shadow](shape-shadow-property-project.md)|Gets or sets the shadow formatting properties for the shape. Read-only  **[ShadowFormat](http://msdn.microsoft.com/en-us/library/office/ff195339%28v=office.15%29)**.|
|[ShapeStyle](shape-shapestyle-property-project.md)|Gets or sets the style of the shape. Read/write  **[MsoShapeStyleIndex](http://msdn.microsoft.com/en-us/library/office/ff862067%28v=office.15%29)**.|
|[SoftEdge](shape-softedge-property-project.md)|Gets soft edge formatting properties for the shape. Read-only  **[SoftEdgeFormat](http://msdn.microsoft.com/en-us/library/office/ff863361%28v=office.15%29)**.|
|[Table](shape-table-property-project.md)|Gets the  **ReportTable** object in the shape. Read-only[ReportTable](reporttable-object-project.md).|
|[TextEffect](shape-texteffect-property-project.md)|Gets text formatting properties for the shape. Read-only  **[TextEffectFormat](http://msdn.microsoft.com/en-us/library/office/ff834714%28v=office.15%29)**.|
|[TextFrame](shape-textframe-property-project.md)|Gets a  **TextFrame** object that contains the alignment and anchoring properties of the shape. Read-only **[TextFrame](http://msdn.microsoft.com/en-us/library/office/ff197860%28v=office.15%29)**.|
|[TextFrame2](shape-textframe2-property-project.md)|Gets a  **TextFrame2** object that contains the text in a text frame and the members that control the alignment, anchoring, and other features of the text frame. Read-only **[TextFrame2](http://msdn.microsoft.com/en-us/library/office/ff822136%28v=office.15%29)**.|
|[ThreeD](shape-threed-property-project.md)|Gets a  **ThreeDFormat** object that contains three-dimensional formatting properties for the shape range. Read-only **[ThreeDFormat](http://msdn.microsoft.com/en-us/library/office/ff836783%28v=office.15%29)**.|
|[Title](shape-title-property-project.md)|Gets or sets the title of the shape. Read/write  **String**.|
|[Top](shape-top-property-project.md)|Gets or sets the vertical distance, in points, from the top of the report pane to the top edge of the shape. Read-write  **Single**.|
|[Type](shape-type-property-project.md)|Gets the shape type. Read-only  **[MsoShapeType](http://msdn.microsoft.com/en-us/library/office/ff860759%28v=office.15%29)**.|
|[VerticalFlip](shape-verticalflip-property-project.md)|Gets a value that indicates whether the shape is flipped around the vertical axis. Read-only  **[MsoTriState](http://msdn.microsoft.com/en-us/library/office/ff860737%28v=office.15%29)**.|
|[Vertices](shape-vertices-property-project.md)|Gets the coordinates of the vertices (and control points for a Bézier curve) as a series of coordinate pairs, for a shape that is a drawing. Read-only  **Variant**.|
|[Visible](shape-visible-property-project.md)|Gets or sets a value that determines whether the shape is visible. Read/write  **[MsoTriState](http://msdn.microsoft.com/en-us/library/office/ff860737%28v=office.15%29)**.|
|[Width](shape-width-property-project.md)|Gets or sets the width, in points, of the shape. Read/write  **Long**.|
|[ZOrderPosition](shape-zorderposition-property-project.md)|Gets the position of the shape in the z-order. Read-only  **Long**.|

