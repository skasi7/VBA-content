---
title: ShapeRange Object (PowerPoint)
keywords: vbapp10.chm548000
f1_keywords:
- vbapp10.chm548000
ms.prod: POWERPOINT
ms.assetid: 0a194183-380e-ffb6-9336-b5bd311e917d
---


# ShapeRange Object (PowerPoint)

Represents a shape range, which is a set of shapes on a document. A shape range can contain as few as a single shape or as many as all the shapes on the document.


## Remarks

You can include whichever shapes you want — chosen from among all the shapes on the document or all the shapes in the selection — to construct a shape range. For example, you could construct a  **ShapeRange** collection that contains the first three shapes on a document, all the selected shapes on a document, or all the freeforms on a document.

For an overview of how to work with either a single shape or with more than one shape at a time, see [How to: Work with Shapes (Drawing Objects)](http://msdn.microsoft.com/library/work-with-shapes-drawing-objects%28Office.15%29.aspx).

The following examples describe how to:


- Return a set of shapes you specify by name or index number.
    
- Return all or some of the selected shapes on a document.
    

## Example

Use  **Shapes.Range** (index), where index is the name or index number of the shape or an array that contains either names or index numbers of shapes, to return a **ShapeRange** collection that represents a set of shapes on a document. You can use the **Array** function to construct an array of names or index numbers. The following example sets the fill pattern for shapes one and three on `myDocument`.


```
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes.Range(Array(1, 3)).Fill _

    .Patterned msoPatternHorizontalBrick
```

The following example sets the fill pattern for the shapes named "Oval 4" and "Rectangle 5" on  `myDocument`.




```
Set myDocument = ActivePresentation.Slides(1)

Set myRange = myDocument.Shapes _

    .Range(Array("Oval 4", "Rectangle 5"))

myRange.Fill.Patterned msoPatternHorizontalBrick
```

Although you can use the [Range](http://msdn.microsoft.com/library/shapes-range-method-powerpoint%28Office.15%29.aspx)method to return any number of shapes or slides, it is simpler to use the [Item](http://msdn.microsoft.com/library/shaperange-item-method-powerpoint%28Office.15%29.aspx)method if you want to return only a single member of the collection. For example,  `Shapes(1)` is simpler than `Shapes.Range(1)`.

Use the [ShapeRange](http://msdn.microsoft.com/library/selection-shaperange-property-powerpoint%28Office.15%29.aspx)property of the  **Selection** object to return all the shapes in the selection. The following example sets the fill foreground color for all the shapes in the selection in window one, assuming that there's at least one shape in the selection.




```
Windows(1).Selection.ShapeRange.Fill.ForeColor _

    .RGB = RGB(255, 0, 255)
```

Use  **Selection.ShapeRange** (index), where index is the shape name or the index number, to return a single shape within the selection. The following example sets the fill foreground color for shape two in the collection of selected shapes in window one, assuming that there are at least two shapes in the selection.




```
Windows(1).Selection.ShapeRange(2).Fill.ForeColor _

    .RGB = RGB(255, 0, 255)
```


## Methods



|**Name**|
|:-----|
|[Align](http://msdn.microsoft.com/library/shaperange-align-method-powerpoint%28Office.15%29.aspx)|
|[Apply](http://msdn.microsoft.com/library/shaperange-apply-method-powerpoint%28Office.15%29.aspx)|
|[ApplyAnimation](http://msdn.microsoft.com/library/shaperange-applyanimation-method-powerpoint%28Office.15%29.aspx)|
|[ConvertTextToSmartArt](http://msdn.microsoft.com/library/shaperange-converttexttosmartart-method-powerpoint%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/shaperange-copy-method-powerpoint%28Office.15%29.aspx)|
|[Cut](http://msdn.microsoft.com/library/shaperange-cut-method-powerpoint%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/shaperange-delete-method-powerpoint%28Office.15%29.aspx)|
|[Distribute](http://msdn.microsoft.com/library/shaperange-distribute-method-powerpoint%28Office.15%29.aspx)|
|[Duplicate](http://msdn.microsoft.com/library/shaperange-duplicate-method-powerpoint%28Office.15%29.aspx)|
|[Flip](http://msdn.microsoft.com/library/shaperange-flip-method-powerpoint%28Office.15%29.aspx)|
|[Group](http://msdn.microsoft.com/library/shaperange-group-method-powerpoint%28Office.15%29.aspx)|
|[IncrementLeft](http://msdn.microsoft.com/library/shaperange-incrementleft-method-powerpoint%28Office.15%29.aspx)|
|[IncrementRotation](http://msdn.microsoft.com/library/shaperange-incrementrotation-method-powerpoint%28Office.15%29.aspx)|
|[IncrementTop](http://msdn.microsoft.com/library/shaperange-incrementtop-method-powerpoint%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/shaperange-item-method-powerpoint%28Office.15%29.aspx)|
|[MergeShapes](http://msdn.microsoft.com/library/shaperange-mergeshapes-method-powerpoint%28Office.15%29.aspx)|
|[PickUp](http://msdn.microsoft.com/library/shaperange-pickup-method-powerpoint%28Office.15%29.aspx)|
|[PickupAnimation](http://msdn.microsoft.com/library/shaperange-pickupanimation-method-powerpoint%28Office.15%29.aspx)|
|[Regroup](http://msdn.microsoft.com/library/shaperange-regroup-method-powerpoint%28Office.15%29.aspx)|
|[RerouteConnections](http://msdn.microsoft.com/library/shaperange-rerouteconnections-method-powerpoint%28Office.15%29.aspx)|
|[ScaleHeight](http://msdn.microsoft.com/library/shaperange-scaleheight-method-powerpoint%28Office.15%29.aspx)|
|[ScaleWidth](http://msdn.microsoft.com/library/shaperange-scalewidth-method-powerpoint%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/shaperange-select-method-powerpoint%28Office.15%29.aspx)|
|[SetShapesDefaultProperties](http://msdn.microsoft.com/library/shaperange-setshapesdefaultproperties-method-powerpoint%28Office.15%29.aspx)|
|[Ungroup](http://msdn.microsoft.com/library/shaperange-ungroup-method-powerpoint%28Office.15%29.aspx)|
|[UpgradeMedia](http://msdn.microsoft.com/library/shaperange-upgrademedia-method-powerpoint%28Office.15%29.aspx)|
|[ZOrder](http://msdn.microsoft.com/library/shaperange-zorder-method-powerpoint%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[ActionSettings](http://msdn.microsoft.com/library/shaperange-actionsettings-property-powerpoint%28Office.15%29.aspx)|
|[Adjustments](http://msdn.microsoft.com/library/shaperange-adjustments-property-powerpoint%28Office.15%29.aspx)|
|[AlternativeText](http://msdn.microsoft.com/library/shaperange-alternativetext-property-powerpoint%28Office.15%29.aspx)|
|[AnimationSettings](http://msdn.microsoft.com/library/shaperange-animationsettings-property-powerpoint%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/shaperange-application-property-powerpoint%28Office.15%29.aspx)|
|[AutoShapeType](http://msdn.microsoft.com/library/shaperange-autoshapetype-property-powerpoint%28Office.15%29.aspx)|
|[BackgroundStyle](http://msdn.microsoft.com/library/shaperange-backgroundstyle-property-powerpoint%28Office.15%29.aspx)|
|[BlackWhiteMode](http://msdn.microsoft.com/library/shaperange-blackwhitemode-property-powerpoint%28Office.15%29.aspx)|
|[Callout](http://msdn.microsoft.com/library/shaperange-callout-property-powerpoint%28Office.15%29.aspx)|
|[Chart](http://msdn.microsoft.com/library/shaperange-chart-property-powerpoint%28Office.15%29.aspx)|
|[Child](http://msdn.microsoft.com/library/shaperange-child-property-powerpoint%28Office.15%29.aspx)|
|[ConnectionSiteCount](http://msdn.microsoft.com/library/shaperange-connectionsitecount-property-powerpoint%28Office.15%29.aspx)|
|[Connector](http://msdn.microsoft.com/library/shaperange-connector-property-powerpoint%28Office.15%29.aspx)|
|[ConnectorFormat](http://msdn.microsoft.com/library/shaperange-connectorformat-property-powerpoint%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/shaperange-count-property-powerpoint%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/shaperange-creator-property-powerpoint%28Office.15%29.aspx)|
|[CustomerData](http://msdn.microsoft.com/library/shaperange-customerdata-property-powerpoint%28Office.15%29.aspx)|
|[Fill](http://msdn.microsoft.com/library/shaperange-fill-property-powerpoint%28Office.15%29.aspx)|
|[Glow](http://msdn.microsoft.com/library/shaperange-glow-property-powerpoint%28Office.15%29.aspx)|
|[GroupItems](http://msdn.microsoft.com/library/shaperange-groupitems-property-powerpoint%28Office.15%29.aspx)|
|[HasChart](http://msdn.microsoft.com/library/shaperange-haschart-property-powerpoint%28Office.15%29.aspx)|
|[HasInkXML](http://msdn.microsoft.com/library/shaperange-hasinkxml-property-powerpoint%28Office.15%29.aspx)|
|[HasSmartArt](http://msdn.microsoft.com/library/shaperange-hassmartart-property-powerpoint%28Office.15%29.aspx)|
|[HasTable](http://msdn.microsoft.com/library/shaperange-hastable-property-powerpoint%28Office.15%29.aspx)|
|[HasTextFrame](http://msdn.microsoft.com/library/shaperange-hastextframe-property-powerpoint%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/shaperange-height-property-powerpoint%28Office.15%29.aspx)|
|[HorizontalFlip](http://msdn.microsoft.com/library/shaperange-horizontalflip-property-powerpoint%28Office.15%29.aspx)|
|[Id](http://msdn.microsoft.com/library/shaperange-id-property-powerpoint%28Office.15%29.aspx)|
|[InkXML](http://msdn.microsoft.com/library/shaperange-inkxml-property-powerpoint%28Office.15%29.aspx)|
|[IsNarration](http://msdn.microsoft.com/library/shaperange-isnarration-property-powerpoint%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/shaperange-left-property-powerpoint%28Office.15%29.aspx)|
|[Line](http://msdn.microsoft.com/library/shaperange-line-property-powerpoint%28Office.15%29.aspx)|
|[LinkFormat](http://msdn.microsoft.com/library/shaperange-linkformat-property-powerpoint%28Office.15%29.aspx)|
|[LockAspectRatio](http://msdn.microsoft.com/library/shaperange-lockaspectratio-property-powerpoint%28Office.15%29.aspx)|
|[MediaFormat](http://msdn.microsoft.com/library/shaperange-mediaformat-property-powerpoint%28Office.15%29.aspx)|
|[MediaType](http://msdn.microsoft.com/library/shaperange-mediatype-property-powerpoint%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/shaperange-name-property-powerpoint%28Office.15%29.aspx)|
|[Nodes](http://msdn.microsoft.com/library/shaperange-nodes-property-powerpoint%28Office.15%29.aspx)|
|[OLEFormat](http://msdn.microsoft.com/library/shaperange-oleformat-property-powerpoint%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/shaperange-parent-property-powerpoint%28Office.15%29.aspx)|
|[ParentGroup](http://msdn.microsoft.com/library/shaperange-parentgroup-property-powerpoint%28Office.15%29.aspx)|
|[PictureFormat](http://msdn.microsoft.com/library/shaperange-pictureformat-property-powerpoint%28Office.15%29.aspx)|
|[PlaceholderFormat](http://msdn.microsoft.com/library/shaperange-placeholderformat-property-powerpoint%28Office.15%29.aspx)|
|[Reflection](http://msdn.microsoft.com/library/shaperange-reflection-property-powerpoint%28Office.15%29.aspx)|
|[Rotation](http://msdn.microsoft.com/library/shaperange-rotation-property-powerpoint%28Office.15%29.aspx)|
|[Shadow](http://msdn.microsoft.com/library/shaperange-shadow-property-powerpoint%28Office.15%29.aspx)|
|[ShapeStyle](http://msdn.microsoft.com/library/shaperange-shapestyle-property-powerpoint%28Office.15%29.aspx)|
|[SmartArt](http://msdn.microsoft.com/library/shaperange-smartart-property-powerpoint%28Office.15%29.aspx)|
|[SoftEdge](http://msdn.microsoft.com/library/shaperange-softedge-property-powerpoint%28Office.15%29.aspx)|
|[Table](http://msdn.microsoft.com/library/shaperange-table-property-powerpoint%28Office.15%29.aspx)|
|[Tags](http://msdn.microsoft.com/library/shaperange-tags-property-powerpoint%28Office.15%29.aspx)|
|[TextEffect](http://msdn.microsoft.com/library/shaperange-texteffect-property-powerpoint%28Office.15%29.aspx)|
|[TextFrame](http://msdn.microsoft.com/library/shaperange-textframe-property-powerpoint%28Office.15%29.aspx)|
|[TextFrame2](http://msdn.microsoft.com/library/shaperange-textframe2-property-powerpoint%28Office.15%29.aspx)|
|[ThreeD](http://msdn.microsoft.com/library/shaperange-threed-property-powerpoint%28Office.15%29.aspx)|
|[Title](http://msdn.microsoft.com/library/shaperange-title-property-powerpoint%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/shaperange-top-property-powerpoint%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/shaperange-type-property-powerpoint%28Office.15%29.aspx)|
|[VerticalFlip](http://msdn.microsoft.com/library/shaperange-verticalflip-property-powerpoint%28Office.15%29.aspx)|
|[Vertices](http://msdn.microsoft.com/library/shaperange-vertices-property-powerpoint%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/shaperange-visible-property-powerpoint%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/shaperange-width-property-powerpoint%28Office.15%29.aspx)|
|[ZOrderPosition](http://msdn.microsoft.com/library/shaperange-zorderposition-property-powerpoint%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/object-model-powerpoint-vba-reference%28Office.15%29.aspx)
