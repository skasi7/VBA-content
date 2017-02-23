---
title: Shape Properties (Excel)
ms.prod: EXCEL
ms.assetid: 4e220346-5175-4b00-ae82-3caaf2f004fb
---


# Shape Properties (Excel)

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Adjustments](shape-adjustments-property-excel.md)|Returns an  **[Adjustments](adjustments-object-excel.md)** object that contains adjustment values for all the adjustments in the specified shape. Applies to any **[Shape](shape-object-excel.md)** object that represents an AutoShape, WordArt, or a connector.|
|[AlternativeText](shape-alternativetext-property-excel.md)|Returns or sets the descriptive (alternative) text string for a  **[Shape](shape-object-excel.md)** object when the object is saved to a Web page. Read/write **String** .|
|[Application](shape-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[AutoShapeType](shape-autoshapetype-property-excel.md)|Returns or sets the shape type for the specified  **[Shape](shape-object-excel.md)** or **[ShapeRange](shaperange-object-excel.md)** object, which must represent an AutoShape other than a line, freeform drawing, or connector. Read/write **[MsoAutoShapeType](msoautoshapetype-enumeration-office.md)** .|
|[BackgroundStyle](shape-backgroundstyle-property-excel.md)|Returns or sets the background style. Read/write  **[MsoBackgroundStyleIndex](msobackgroundstyleindex-enumeration-office.md)** .|
|[BlackWhiteMode](shape-blackwhitemode-property-excel.md)|Returns or sets a value that indicates how the specified shape appears when the presentation is viewed in black-and-white mode. Read/write  **[MsoBlackWhiteMode](msoblackwhitemode-enumeration-office.md)** .|
|[BottomRightCell](shape-bottomrightcell-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents the cell that lies under the lower-right corner of the object. Read-only.|
|[Callout](shape-callout-property-excel.md)|Returns a  **[CalloutFormat](calloutformat-object-excel.md)** object that contains callout formatting properties for the specified shape. Applies to a **[Shape](shape-object-excel.md)** object that represent line callouts. Read-only.|
|[Chart](shape-chart-property-excel.md)|Returns a  **[Chart](chart-object-excel.md)** object that represents the chart contained in the shape. Read-only.|
|[Child](shape-child-property-excel.md)|Returns  **msoTrue** if the specified shape is a child shape or if all shapes in a shape range are child shapes of the same parent. Read-only **[MsoTriState](msotristate-enumeration-office.md)** .|
|[ConnectionSiteCount](shape-connectionsitecount-property-excel.md)|Returns the number of connection sites on the specified shape. Read-only  **Long** .|
|[Connector](shape-connector-property-excel.md)| **True** if the specified shape is a connector. Read-only **[MsoTriState](msotristate-enumeration-office.md)** .|
|[ConnectorFormat](shape-connectorformat-property-excel.md)|Returns a  **[ConnectorFormat](connectorformat-object-excel.md)** object that contains connector formatting properties. Applies to a **[Shape](shape-object-excel.md)** that represent connectors. Read-only.|
|[ControlFormat](shape-controlformat-property-excel.md)|Returns a  **[ControlFormat](controlformat-object-excel.md)** object that contains Microsoft Excel control properties. Read-only.|
|[Creator](shape-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[Fill](shape-fill-property-excel.md)|Returns a  **[FillFormat](fillformat-object-excel.md)** object for a specified shape or a **[ChartFillFormat](chartfillformat-object.md)** object for a specified chart that contains fill formatting properties for the shape or chart. Read-only.|
|[FormControlType](shape-formcontroltype-property-excel.md)|Returns the Microsoft Excel control type. Read-only  **[XlFormControl](xlformcontrol-enumeration-excel.md)** .|
|[Glow](shape-glow-property-excel.md)|Returns a  **[GlowFormat](glowformat-object-office.md)** object for a specified shape that contains glow formatting properties for the shape. Read-only.|
|[GroupItems](shape-groupitems-property-excel.md)|Returns a  **[GroupShapes](groupshapes-object-excel.md)** object that represents the individual shapes in the specified group. Use the **[Item](groupshapes-item-method-excel.md)** method of the **GroupShapes** object to return a single shape from the group. Applies to **Shape** objects that represent grouped shapes. Read-only.|
|[HasChart](shape-haschart-property-excel.md)| Returns whether a shape contains a chart. Read-only **[MsoTriState](msotristate-enumeration-office.md)** .|
|[HasSmartArt](shape-hassmartart-property-excel.md)|Returns whether there is a SmartArt diagram present on the specified shape. Read-only|
|[Height](shape-height-property-excel.md)|Returns or sets a  **Single** value that represents the height, in points, of the object.|
|[HorizontalFlip](shape-horizontalflip-property-excel.md)| **True** if the specified shape is flipped around the horizontal axis. Read-only **[MsoTriState](msotristate-enumeration-office.md)** .|
|[Hyperlink](shape-hyperlink-property-excel.md)|Returns a  **[Hyperlink](hyperlink-object-excel.md)** object that represents the hyperlink for the shape.|
|[ID](shape-id-property-excel.md)|Returns a Long value that represents the type for the specified object.|
|[Left](shape-left-property-excel.md)|Returns or sets a  **Single** value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).|
|[Line](shape-line-property-excel.md)|Returns a  **[LineFormat](lineformat-object-excel.md)** object that contains line formatting properties for the specified shape. (For a line, the **LineFormat** object represents the line itself; for a shape with a border, the **LineFormat** object represents the border). Read-only.|
|[LinkFormat](shape-linkformat-property-excel.md)|Returns a  **[LinkFormat](linkformat-object-excel.md)** object that contains linked OLE object properties. Read-only.|
|[LockAspectRatio](shape-lockaspectratio-property-excel.md)| **True** if the specified shape retains its original proportions when you resize it. **False** if you can change the height and width of the shape independently of one another when you resize it. Read/write **[MsoTriState](msotristate-enumeration-office.md)** .|
|[Locked](shape-locked-property-excel.md)|Returns or sets a  **Boolean** value that indicates if the object is locked.|
|[Name](shape-name-property-excel.md)|Returns or sets a  **String** value representing the name of the object.|
|[Nodes](shape-nodes-property-excel.md)|Returns a  **[ShapeNodes](shapenodes-object-excel.md)** collection that represents the geometric description of the specified shape.|
|[OLEFormat](shape-oleformat-property-excel.md)|Returns an  **[OLEFormat](oleformat-object-excel.md)** object that contains OLE object properties. Read-only.|
|[OnAction](shape-onaction-property-excel.md)|Returns or sets the name of a macro that's run when the specified object is clicked. Read/write  **String** .|
|[Parent](shape-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[ParentGroup](shape-parentgroup-property-excel.md)|Returns a  **[Shape](shape-object-excel.md)** object that represents the common parent shape of a child shape or a range of child shapes.|
|[PictureFormat](shape-pictureformat-property-excel.md)|Returns a  **[PictureFormat](pictureformat-object-excel.md)** object that contains picture formatting properties for the specified shape. Applies to a **[Shape](shape-object-excel.md)** object that represent pictures or OLE objects. Read-only.|
|[Placement](shape-placement-property-excel.md)|Returns or sets an  **[XlPlacement](xlplacement-enumeration-excel.md)** value that represents the way the object is attached to the cells below it.|
|[Reflection](shape-reflection-property-excel.md)|Returns a  **[ReflectionFormat](reflectionformat-object-office.md)** object for a specified shape that contains reflection formatting properties for the shape. Read-only.|
|[Rotation](shape-rotation-property-excel.md)|Returns or sets the rotation of the shape, in degrees. Read/write  **Single** .|
|[Shadow](shape-shadow-property-excel.md)|Returns a read-only  **[ShadowFormat](shadowformat-object-excel.md)** object that contains shadow formatting properties for the specified shape or shapes.|
|[ShapeStyle](shape-shapestyle-property-excel.md)|Returns or sets an  **[MsoShapeStyleIndex](msoshapestyleindex-enumeration-office.md)** that represents the shape style of shape range. Read/write.|
|[SmartArt](shape-smartart-property-excel.md)|Returns an object that represents the SmartArt associated with the shape. Read-only|
|[SoftEdge](shape-softedge-property-excel.md)|Returns a  **[SoftEdgeFormat](softedgeformat-object-office.md)** object for a specified shape that contains soft edge formatting properties for the shape. Read-only.|
|[TextEffect](shape-texteffect-property-excel.md)|Returns a  **[TextEffectFormat](texteffectformat-object-excel.md)** object that contains text-effect formatting properties for the specified shape. Read-only.|
|[TextFrame](shape-textframe-property-excel.md)|Returns a  **[TextFrame](textframe-object-excel.md)** object that contains the alignment and anchoring properties for the specified shape. Read-only.|
|[TextFrame2](shape-textframe2-property-excel.md)|Returns a  **[TextFrame2](textframe2-object-excel.md)** object that contains text formatting for the specified shape. Read-only.|
|[ThreeD](shape-threed-property-excel.md)|Returns a  **[ThreeDFormat](threedformat-object-excel.md)** object that contains 3-D - effect formatting properties for the specified shape. Read-only.|
|[Title](shape-title-property-excel.md)|Returns or sets the title of the alternative text associated with the specified shape. Read/write|
|[Top](shape-top-property-excel.md)|Returns or sets a  **Single** value that represents the distance, in points, from the top edge of the topmost shape in the shape range to the top edge of the worksheet.|
|[TopLeftCell](shape-topleftcell-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents the cell that lies under the upper-left corner of the specified object. Read-only.|
|[Type](shape-type-property-excel.md)|Returns or sets a  **[MsoShapeType](msoshapetype-enumeration-office.md)** value that represents the shape type.|
|[VerticalFlip](shape-verticalflip-property-excel.md)| **True** if the specified shape is flipped around the vertical axis. Read-only **[MsoTriState](msotristate-enumeration-office.md)** .|
|[Vertices](shape-vertices-property-excel.md)|Returns the coordinates of the specified freeform drawing's vertices (and control points for Bézier curves) as a series of coordinate pairs. You can use the array returned by this property as an argument to the  **[AddCurve](shapes-addcurve-method-excel.md)** method or **[AddPolyLine](shapes-addpolyline-method-excel.md)** method. Read-only **Variant** .|
|[Visible](shape-visible-property-excel.md)|Returns or sets a  **[MsoTriState](msotristate-enumeration-office.md)** value that determines whether the object is visible. Read/write.|
|[Width](shape-width-property-excel.md)|Returns or sets a  **Single** value that represents the width, in points, of the object.|
|[ZOrderPosition](shape-zorderposition-property-excel.md)|Returns the position of the specified shape in the z-order. Read-only  **Long** .Read-only|

