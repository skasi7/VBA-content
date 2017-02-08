---
title: ShapeRange Members (Project)
ms.prod: PROJECTSERVER
ms.assetid: b69ad896-7b8e-4402-a191-60eb321d4bfd
---


# ShapeRange Members (Project)
Represents a shape range, which is a collection of one or more shapes in a report.

Represents a shape range, which is a collection of one or more shapes in a report.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Align](shaperange-align-method-project.md)|The  **Align** method is not implemented in Project.|
|[Apply](shaperange-apply-method-project.md)|Applies formatting to a shape range, where the formatting information has been copied by using the  **[PickUp](shape-pickup-method-project.md)** method.|
|[Copy](shaperange-copy-method-project.md)|Copies the shape range to the Clipboard.|
|[Cut](shaperange-cut-method-project.md)|Cuts the shape range to the Clipboard.|
|[Delete](shaperange-delete-method-project.md)|Deletes the shape range.|
|[Distribute](shaperange-distribute-method-project.md)|The  **Distribute** method is not implemented in Project.|
|[Duplicate](shaperange-duplicate-method-project.md)|Duplicates a shape range and returns a reference to the copy.|
|[Flip](shaperange-flip-method-project.md)|Flips each shape in the shape range around its horizontal or vertical axis.|
|[Group](shaperange-group-method-project.md)|The  **Group** method is not implemented in Project.|
|[IncrementLeft](shaperange-incrementleft-method-project.md)|Moves each shape in the shape range horizontally by the specified number of points.|
|[IncrementRotation](shaperange-incrementrotation-method-project.md)|Rotates each shape in the shape range around the z-axis by the specified number of degrees.|
|[IncrementTop](shaperange-incrementtop-method-project.md)|Moves each shape in the shape range vertically by the specified number of points.|
|[Item](shaperange-item-method-project.md)|Gets an individual  **Shape** object in the shape range collection.|
|[MergeShapes](shaperange-mergeshapes-method-project.md)|The  **MergeShapes** method is not implemented in Project.|
|[PickUp](shaperange-pickup-method-project.md)|Copies the formatting of the shape range.|
|[Regroup](shaperange-regroup-method-project.md)|The  **Regroup** method is not implemented in Project.|
|[RerouteConnections](shaperange-rerouteconnections-method-project.md)|The  **RerouteConnections** method is not implemented in Project.|
|[ScaleHeight](shaperange-scaleheight-method-project.md)|Scales the height of the range of shapes by a specified factor.|
|[ScaleWidth](shaperange-scalewidth-method-project.md)|Scales the width of the range of shapes by a specified factor.|
|[Select](shaperange-select-method-project.md)|Selects each shape in a shape range.|
|[SetShapesDefaultProperties](shaperange-setshapesdefaultproperties-method-project.md)|Applies the formatting of a default shape to each shape in the range.|
|[Ungroup](shaperange-ungroup-method-project.md)|The  **Ungroup** method is not implemented in Project.|
|[ZOrder](shaperange-zorder-method-project.md)|Moves the shape range in front of or behind other shapes (that is, changes the position in the z-order).|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Adjustments](shaperange-adjustments-property-project.md)|Gets an  **[Adjustments](http://msdn.microsoft.com/en-us/library/office/ff838852%28v=office.15%29)** object that contains adjustment values for all the adjustments in the shape. Applies to any **ShapeRange** object that represents an AutoShape, WordArt, or a connector. Read-only **Adjustments**.|
|[AlternativeText](shaperange-alternativetext-property-project.md)|Gets or sets the descriptive (alternative) text string for a  **ShapeRange** object when the object is saved to a web page. Read/write **String**.|
|[Application](shaperange-application-property-project.md)|Gets the [Application Object (Project)](application-object-project.md) object. Read-only **Application**.|
|[AutoShapeType](shaperange-autoshapetype-property-project.md)|Gets or sets the shape type for the  **ShapeRange** object, which must represent an AutoShape other than a line, freeform drawing, or connector. Read/write **MsoAutoShapeType**.|
|[BackgroundStyle](shaperange-backgroundstyle-property-project.md)|Gets or sets the background style. Read/write  **MsoBackgroundStyleIndex**.|
|[BlackWhiteMode](shaperange-blackwhitemode-property-project.md)|Gets or sets a value that indicates how the shape appears when it is viewed in black-and-white mode. Read/write  **MsoBlackWhiteMode**.|
|[Callout](shaperange-callout-property-project.md)|Gets callout formatting properties for the shape range, when the  **ShapeRange** object represents a callout. Read-only **CalloutFormat**.|
|[Chart](shaperange-chart-property-project.md)|Gets a  **Chart** object that represents the chart contained in the shape range. Read-only **Chart**.|
|[Child](shaperange-child-property-project.md)|Gets a value that indicates whether all shapes in the shape range are child shapes of the same parent. Read-only  **[MsoTriState](http://msdn.microsoft.com/en-us/library/office/ff860737%28v=office.15%29)**.|
|[ConnectionSiteCount](shaperange-connectionsitecount-property-project.md)|Gets the number of connection sites on the shape range. Read-only  **Long**.|
|[Connector](shaperange-connector-property-project.md)|Gets a value that indicates whether the shape range is a connector. Read-only  **[MsoTriState](http://msdn.microsoft.com/en-us/library/office/ff860737%28v=office.15%29)**|
|[ConnectorFormat](shaperange-connectorformat-property-project.md)|Gets a  **ConnectorFormat** object that contains connector formatting properties. Applies to a **ShapeRange** object that represents one or more connectors. Read-only **[ConnectorFormat](http://msdn.microsoft.com/en-us/library/office/ff820940%28v=office.15%29)**.|
|[Count](shaperange-count-property-project.md)|Gets the number of shapes in the shape range. Read-only  **Long**.|
|[Fill](shaperange-fill-property-project.md)|Gets a  **FillFormat** object for the shape range, if the shape range contains fill formatting properties. Read-only **[FillFormat](http://msdn.microsoft.com/en-us/library/office/ff838198%28v=office.15%29)**.|
|[Glow](shaperange-glow-property-project.md)|Gets a  **GlowFormat** object for the shape range, if the shape range contains glow formatting properties. Read-only **[GlowFormat](http://msdn.microsoft.com/en-us/library/office/ff864010%28v=office.15%29)**.|
|[GroupItems](shaperange-groupitems-property-project.md)|Gets a  **GroupShapes** object that represents the individual shapes in a group, if the **ShapeRange** object represents a group of shapes. Read-only[GroupShapes](http://msdn.microsoft.com/en-us/library/office/ff195331%28v=office.15%29).|
|[HasChart](shaperange-haschart-property-project.md)|Gets a value that indicates whether the shape range contains a chart. Read-only  **[MsoTriState](http://msdn.microsoft.com/en-us/library/office/ff860737%28v=office.15%29)**.|
|[HasTable](shaperange-hastable-property-project.md)|Gets a value that indicates whether the shape range contains a table. Read-only  **[MsoTriState](http://msdn.microsoft.com/en-us/library/office/ff860737%28v=office.15%29)**.|
|[Height](shaperange-height-property-project.md)|Gets or sets the height of the shape range, in points. Read-write  **Single**.|
|[HorizontalFlip](shaperange-horizontalflip-property-project.md)|Gets a value that indicates whether the shape range is flipped around the horizontal axis. Read-only  **[MsoTriState](http://msdn.microsoft.com/en-us/library/office/ff860737%28v=office.15%29)**.|
|[ID](shaperange-id-property-project.md)|Gets the identification type for the shape range. Read-only  **Long**.|
|[Left](shaperange-left-property-project.md)|Gets or sets the horizontal distance, in points, from the left side of the report to the left edge of the shape range. Read-write  **Single**.|
|[Line](shaperange-line-property-project.md)|Gets the line formatting properties for the shape range. Read-only  **[LineFormat](http://msdn.microsoft.com/en-us/library/office/ff194214%28v=office.15%29)**.|
|[LockAspectRatio](shaperange-lockaspectratio-property-project.md)|Gets or sets a value that indicates whether the shape range retains its original proportions when you resize it; that is, the aspect ratio of the shape range is locked. Read-write  **[MsoTriState](http://msdn.microsoft.com/en-us/library/office/ff860737%28v=office.15%29)**.|
|[Name](shaperange-name-property-project.md)|Gets or sets the name of the shape range. Read-write  **String**.|
|[Nodes](shaperange-nodes-property-project.md)|Gets the geometric description of nodes or control points in the shape range. Read-only  **[ShapeNodes](http://msdn.microsoft.com/en-us/library/office/ff822109%28v=office.15%29)**.|
|[Parent](shaperange-parent-property-project.md)|Gets the parent object of the shape. Read-only  **Object**.|
|[ParentGroup](shaperange-parentgroup-property-project.md)|Gets the common parent shape of a child shape or a range of child shapes. Read-only  **Shape**.|
|[Reflection](shaperange-reflection-property-project.md)|Gets the reflection formatting for the shape range. Read-only  **[ReflectionFormat](http://msdn.microsoft.com/en-us/library/office/ff863140%28v=office.15%29)**.|
|[Rotation](shaperange-rotation-property-project.md)|Gets or sets the rotation of the shape range, in degrees. Read/write  **Single**.|
|[Script](shaperange-script-property-project.md)||
|[Shadow](shaperange-shadow-property-project.md)|Gets or sets the shadow formatting properties for the shape range. Read-only  **[ShadowFormat](http://msdn.microsoft.com/en-us/library/office/ff195339%28v=office.15%29)**.|
|[ShapeStyle](shaperange-shapestyle-property-project.md)|Gets or sets the style of the shape range. Read/write  **[MsoShapeStyleIndex](http://msdn.microsoft.com/en-us/library/office/ff862067%28v=office.15%29)**.|
|[SoftEdge](shaperange-softedge-property-project.md)|Gets soft edge formatting properties for the shape range. Read-only  **[SoftEdgeFormat](http://msdn.microsoft.com/en-us/library/office/ff863361%28v=office.15%29)**.|
|[Table](shaperange-table-property-project.md)|Gets the  **ReportTable** object in the shape range. Read-only[ReportTable](reporttable-object-project.md).|
|[TextEffect](shaperange-texteffect-property-project.md)|Gets text formatting properties for the shape range. Read-only  **[TextEffectFormat](http://msdn.microsoft.com/en-us/library/office/ff834714%28v=office.15%29)**.|
|[TextFrame](shaperange-textframe-property-project.md)|Gets a  **TextFrame** object that contains the alignment and anchoring properties of the shape range. Read-only **[TextFrame](http://msdn.microsoft.com/en-us/library/office/ff197860%28v=office.15%29)**.|
|[TextFrame2](shaperange-textframe2-property-project.md)|Gets a  **TextFrame2** object that contains the text in a text frame and the members that control the alignment, anchoring, and other features of the text frame. Read-only **[TextFrame2](http://msdn.microsoft.com/en-us/library/office/ff822136%28v=office.15%29)**.|
|[ThreeD](shaperange-threed-property-project.md)|Gets a  **ThreeDFormat** object that contains 3-D formatting properties for the shape range. Read-only **[ThreeDFormat](http://msdn.microsoft.com/en-us/library/office/ff836783%28v=office.15%29)**.|
|[Title](shaperange-title-property-project.md)|Gets or sets the title of the shapes in the shape range. Read/write  **String**.|
|[Top](shaperange-top-property-project.md)|Gets or sets the vertical distance, in points, from the top of the report pane to the top edge of the shape range. Read-write  **Single**.|
|[Type](shaperange-type-property-project.md)|Gets the the type of shape in the shape range. Read-only  **[MsoShapeType](http://msdn.microsoft.com/en-us/library/office/ff860759%28v=office.15%29)**.|
|[Value](shaperange-value-property-project.md)|Gets an individual  **Shape** object in the **ShapeRange** collection. Read-only **Shape**.|
|[VerticalFlip](shaperange-verticalflip-property-project.md)|Gets a value that indicates whether the shape range is flipped around the vertical axis. Read-only  **[MsoTriState](http://msdn.microsoft.com/en-us/library/office/ff860737%28v=office.15%29)**.|
|[Vertices](shaperange-vertices-property-project.md)|Gets the coordinates of the vertices (and control points for a Bézier curve) as a series of coordinate pairs, for a shape range that contains a drawing. Read-only  **Variant**.|
|[Visible](shaperange-visible-property-project.md)|Gets or sets a value that determines whether all of the shapes in the shape range are visible. Read/write  **[MsoTriState](http://msdn.microsoft.com/en-us/library/office/ff860737%28v=office.15%29)**.|
|[Width](shaperange-width-property-project.md)|Gets or sets the width, in points, of the shapes within the range. Read/write  **Long**.|
|[ZOrderPosition](shaperange-zorderposition-property-project.md)|Gets the position of the shape range in the z-order. Read-only  **Long**.|

