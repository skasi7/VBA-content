---
title: ShapeRange Members (Excel)
ms.prod: EXCEL
ms.assetid: 1d1950c5-32ac-dfc0-8c19-07159a29a2a0
---


# ShapeRange Members (Excel)
Represents a shape range, which is a set of shapes on a document.

Represents a shape range, which is a set of shapes on a document.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Align](shaperange-align-method-excel.md)|Aligns the shapes in the specified range of shapes.|
|[Apply](shaperange-apply-method-excel.md)|Applies to the specified shape formatting that's been copied by using the  **[PickUp](shaperange-pickup-method-excel.md)** method.|
|[Delete](shaperange-delete-method-excel.md)|Deletes the object.|
|[Distribute](shaperange-distribute-method-excel.md)|Horizontally or vertically distributes the shapes in the specified range of shapes.|
|[Duplicate](shaperange-duplicate-method-excel.md)|Duplicates the object and returns a reference to the new copy.|
|[Flip](shaperange-flip-method-excel.md)|Flips the specified shape around its horizontal or vertical axis.|
|[Group](shaperange-group-method-excel.md)|Groups the shapes in the specified range.|
|[IncrementLeft](shaperange-incrementleft-method-excel.md)|Moves the specified shape horizontally by the specified number of points.|
|[IncrementRotation](shaperange-incrementrotation-method-excel.md)|Changes the rotation of the specified shape around the z-axis by the specified number of degrees. Use the  **[Rotation](shaperange-rotation-property-excel.md)** property to set the absolute rotation of the shape.|
|[IncrementTop](shaperange-incrementtop-method-excel.md)|Moves the specified shape vertically by the specified number of points.|
|[Item](shaperange-item-method-excel.md)|Returns a single object from a collection.|
|[PickUp](shaperange-pickup-method-excel.md)|Copies the formatting of the specified shape. Use the  **[Apply](shaperange-apply-method-excel.md)** method to apply the copied formatting to another shape.|
|[Regroup](shaperange-regroup-method-excel.md)|Regroups the group that the specified shape range belonged to previously. Returns the regrouped shapes as a single  **[Shape](shape-object-excel.md)** object.|
|[RerouteConnections](shaperange-rerouteconnections-method-excel.md)|This method reroutes all connectors attached to the specified shape; if the specified shape is a connector, it's rerouted.|
|[ScaleHeight](shaperange-scaleheight-method-excel.md)|Scales the height of the shape by a specified factor. For pictures and OLE objects, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures and OLE objects are always scaled relative to their current height.|
|[ScaleWidth](shaperange-scalewidth-method-excel.md)|Scales the width of the shape by a specified factor. For pictures and OLE objects, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures and OLE objects are always scaled relative to their current width.|
|[Select](shaperange-select-method-excel.md)|Selects the object.|
|[SetShapesDefaultProperties](shaperange-setshapesdefaultproperties-method-excel.md)|Makes the formatting of the specified shape the default formatting for the shape.|
|[Ungroup](shaperange-ungroup-method-excel.md)|Ungroups any grouped shapes in the specified shape or range of shapes. Disassembles pictures and OLE objects within the specified shape or range of shapes.|
|[ZOrder](shaperange-zorder-method-excel.md)|Moves the specified shape in front of or behind other shapes in the collection (that is, changes the shape's position in the z-order).|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Adjustments](shaperange-adjustments-property-excel.md)|Returns an  **[Adjustments](adjustments-object-excel.md)** object that contains adjustment values for all the adjustments in the specified shape. Applies to any **[ShapeRange](shaperange-object-excel.md)** object that represents an AutoShape, WordArt, or a connector.|
|[AlternativeText](shaperange-alternativetext-property-excel.md)|Returns or sets the descriptive (alternative) text string for a  **[ShapeRange](shaperange-object-excel.md)** object when the object is saved to a Web page. Read/write **String** .|
|[Application](shaperange-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[AutoShapeType](shaperange-autoshapetype-property-excel.md)|Returns or sets the shape type for the specified  **[Shape](shape-object-excel.md)** or **[ShapeRange](shaperange-object-excel.md)** object, which must represent an AutoShape other than a line, freeform drawing, or connector. Read/write **[MsoAutoShapeType](msoautoshapetype-enumeration-office.md)** .|
|[BackgroundStyle](shaperange-backgroundstyle-property-excel.md)|Returns or sets the background style. Read/write  **[MsoBackgroundStyleIndex](msobackgroundstyleindex-enumeration-office.md)** .|
|[BlackWhiteMode](shaperange-blackwhitemode-property-excel.md)|Returns or sets a value that indicates how the specified shape appears when the presentation is viewed in black-and-white mode. Read/write  **[MsoBlackWhiteMode](msoblackwhitemode-enumeration-office.md)** .|
|[Callout](shaperange-callout-property-excel.md)|Returns a  **[CalloutFormat](calloutformat-object-excel.md)** object that contains callout formatting properties for the specified shape. Applies to a **[ShapeRange](shaperange-object-excel.md)** object that represent line callouts. Read-only.|
|[Chart](shaperange-chart-property-excel.md)|Returns a  **[Chart](chart-object-excel.md)** object that represents the chart contained in the shape range. Read-only.|
|[Child](shaperange-child-property-excel.md)|Returns  **msoTrue** if the specified shape is a child shape or if all shapes in a shape range are child shapes of the same parent. Read-only **[MsoTriState](msotristate-enumeration-office.md)** .|
|[ConnectionSiteCount](shaperange-connectionsitecount-property-excel.md)|Returns the number of connection sites on the specified shape. Read-only  **Long** .|
|[Connector](shaperange-connector-property-excel.md)| **True** if the specified shape is a connector. Read-only **[MsoTriState](msotristate-enumeration-office.md)** .|
|[ConnectorFormat](shaperange-connectorformat-property-excel.md)|Returns a  **[ConnectorFormat](connectorformat-object-excel.md)** object that contains connector formatting properties. Applies to a **[ShapeRange](shaperange-object-excel.md)** objects that represent connectors. Read-only.|
|[Count](shaperange-count-property-excel.md)|Returns a  **Long** value that represents the number of objects in the collection.|
|[Creator](shaperange-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[Fill](shaperange-fill-property-excel.md)|Returns a  **[FillFormat](fillformat-object-excel.md)** object for a specified shape or a **[ChartFillFormat](chartfillformat-object.md)** object for a specified chart that contains fill formatting properties for the shape or chart. Read-only.|
|[Glow](shaperange-glow-property-excel.md)|Returns a  **[GlowFormat](glowformat-object-office.md)** object for a specified shape range that contains glow formatting properties for the shape range. Read-only.|
|[GroupItems](shaperange-groupitems-property-excel.md)|Returns a  **[GroupShapes](groupshapes-object-excel.md)** object that represents the individual shapes in the specified group. Use the **[Item](groupshapes-item-method-excel.md)** method of the **GroupShapes** object to return a single shape from the group. Applies to **ShapeRange** objects that represent grouped shapes. Read-only.|
|[HasChart](shaperange-haschart-property-excel.md)| Returns whether a shape range contains a chart. Read-only **[MsoTriState](msotristate-enumeration-office.md)** .|
|[Height](shaperange-height-property-excel.md)|Returns or sets a  **Single** value that represents the height, in points, of the object.|
|[HorizontalFlip](shaperange-horizontalflip-property-excel.md)| **True** if the specified shape is flipped around the horizontal axis. Read-only **[MsoTriState](msotristate-enumeration-office.md)** .|
|[ID](shaperange-id-property-excel.md)|Returns a Long value that represents the type for the specified object.|
|[Left](shaperange-left-property-excel.md)|Returns or sets a  **Single** value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).|
|[Line](shaperange-line-property-excel.md)|Returns a  **[LineFormat](lineformat-object-excel.md)** object that contains line formatting properties for the specified shape. (For a line, the **LineFormat** object represents the line itself; for a shape with a border, the **LineFormat** object represents the border). Read-only.|
|[LockAspectRatio](shaperange-lockaspectratio-property-excel.md)| **True** if the specified shape retains its original proportions when you resize it. **False** if you can change the height and width of the shape independently of one another when you resize it. Read/write **[MsoTriState](msotristate-enumeration-office.md)** .|
|[Name](shaperange-name-property-excel.md)|Returns or sets a  **String** value representing the name of the object.|
|[Nodes](shaperange-nodes-property-excel.md)|Returns a  **[ShapeNodes](shapenodes-object-excel.md)** collection that represents the geometric description of the specified shape.|
|[Parent](shaperange-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[ParentGroup](shaperange-parentgroup-property-excel.md)|Returns a  **[Shape](shape-object-excel.md)** object that represents the common parent shape of a child shape or a range of child shapes.|
|[PictureFormat](shaperange-pictureformat-property-excel.md)|Returns a  **[PictureFormat](pictureformat-object-excel.md)** object that contains picture formatting properties for the specified shape. Applies to a **[ShapeRange](shaperange-object-excel.md)** object that represent pictures or OLE objects. Read-only.|
|[Reflection](shaperange-reflection-property-excel.md)|Returns a  **[ReflectionFormat](reflectionformat-object-office.md)** object for a specified shape range that contains reflection formatting properties for the shape range. Read-only.|
|[Rotation](shaperange-rotation-property-excel.md)|Returns or sets the rotation of the shape, in degrees. Read/write  **Single** .|
|[Shadow](shaperange-shadow-property-excel.md)|Returns a read-only  **[ShadowFormat](shadowformat-object-excel.md)** object that contains shadow formatting properties for the specified shape or shapes.|
|[ShapeStyle](shaperange-shapestyle-property-excel.md)|Returns or sets an  **[MsoShapeStyleIndex](msoshapestyleindex-enumeration-office.md)** that represents shape style of shape range. Read/write.|
|[SoftEdge](shaperange-softedge-property-excel.md)|Returns a  **[SoftEdgeFormat](softedgeformat-object-office.md)** object for a specified shape range that contains soft edge formatting properties for the shape range. Read-only.|
|[TextEffect](shaperange-texteffect-property-excel.md)|Returns a  **[TextEffectFormat](texteffectformat-object-excel.md)** object that contains text-effect formatting properties for the specified shape. Read-only.|
|[TextFrame](shaperange-textframe-property-excel.md)|Returns a  **[TextFrame](textframe-object-excel.md)** object that contains the alignment and anchoring properties for the specified shape. Read-only.|
|[TextFrame2](shaperange-textframe2-property-excel.md)|Returns a  **[TextFrame2](textframe2-object-excel.md)** object that contains text formatting for the specified shape range. Read-only.|
|[ThreeD](shaperange-threed-property-excel.md)|Returns a  **[ThreeDFormat](threedformat-object-excel.md)** object that contains 3-D - effect formatting properties for the specified shape. Read-only.|
|[Title](shaperange-title-property-excel.md)|Returns or sets the title of the alternative text associated with the specified shape range. Read/write|
|[Top](shaperange-top-property-excel.md)|Returns or sets a  **Single** value that represents the distance, in points, from the top edge of the topmost shape in the shape range to the top edge of the worksheet.|
|[Type](shaperange-type-property-excel.md)|Returns a  **[MsoShapeType](msoshapetype-enumeration-office.md)** value that represents the shape type.|
|[VerticalFlip](shaperange-verticalflip-property-excel.md)| **True** if the specified shape is flipped around the vertical axis. Read-only **[MsoTriState](msotristate-enumeration-office.md)** .|
|[Vertices](shaperange-vertices-property-excel.md)|Returns the coordinates of the specified freeform drawing's vertices (and control points for Bézier curves) as a series of coordinate pairs. You can use the array returned by this property as an argument to the  **[AddCurve](shapes-addcurve-method-excel.md)** method or **[AddPolyLine](shapes-addpolyline-method-excel.md)** method. Read-only **Variant** .|
|[Visible](shaperange-visible-property-excel.md)|Returns or sets a  **[MsoTriState](msotristate-enumeration-office.md)** value that determines whether the object is visible. Read/write.|
|[Width](shaperange-width-property-excel.md)|Returns or sets a  **Single** value that represents the width, in points, of the object.|
|[ZOrderPosition](shaperange-zorderposition-property-excel.md)|Returns the position of the specified shape in the z-order. Read-only  **Long** .Read-only|

