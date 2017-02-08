---
title: Shape Object (PowerPoint)
keywords: vbapp10.chm547000
f1_keywords:
- vbapp10.chm547000
ms.prod: POWERPOINT
ms.assetid: 1da93849-99e0-827e-ced3-c6cf7f8569f3
---


# Shape Object (PowerPoint)

Represents an object in the drawing layer, such as an AutoShape, freeform, OLE object, or picture.


## Remarks


 **Note**  There are three objects that represent shapes: the  **Shapes** collection, which represents all the shapes on a document; the **[ShapeRange](http://msdn.microsoft.com/library/shaperange-object-powerpoint%28Office.15%29.aspx)** collection, which represents a specified subset of the shapes on a document (for example, a **ShapeRange** object could represent shapes one and four on the document, or it could represent all the selected shapes on the document); the **Shape** object, which represents a single shape on a document. If you want to work with several shape at the same time or with shapes within the selection, use a **ShapeRange** collection. For an overview of how to work with either a single shape or with more than one shape at a time, see [How to: Work with Shapes (Drawing Objects)](http://msdn.microsoft.com/library/work-with-shapes-drawing-objects%28Office.15%29.aspx).

The following examples describe how to:


- Return an existing shape on a slide, indexed by name or number.
    
- Return a newly created shape on a slide.
    
- Return a shape within the selection.
    
- Return the slide title and other placeholders on a slide.
    
- Return the shapes attached to the ends of a connector.
    
- Return the default shape for a presentation.
    
- Return a newly created freeform.
    
- Return a single shape from within a group.
    
- Return a newly formed group of shapes.
    

## Example

Use  **Shapes** (index), where index is the shape name or the index number, to return a **Shape** object that represents a shape on a slide. The following example horizontally flips shape one and the shape named Rectangle 1 on myDocument.


```
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes(1).Flip msoFlipHorizontal

myDocument.Shapes("Rectangle 1").Flip msoFlipHorizontal
```

Each shape is assigned a default name when you add it to the  **Shapes** collection. To give the shape a more meaningful name, use the **Name** property. The following example adds a rectangle to myDocument, gives it the name Red Square, and then sets its foreground color and line style.




```
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes.AddShape(Type:=msoShapeRectangle, _

        Top:=144, Left:=144, Width:=72, Height:=72)

    .Name = "Red Square"

    .Fill.ForeColor.RGB = RGB(255, 0, 0)

    .Line.DashStyle = msoLineDashDot

End With
```

To add a shape to a slide and return a  **Shape** object that represents the newly created shape, use one of the following methods of the **Shapes** collection:[AddCallout](http://msdn.microsoft.com/library/shapes-addcallout-method-powerpoint%28Office.15%29.aspx), [AddComment](http://msdn.microsoft.com/library/11347ca1-cef3-0923-2544-cb80e7fc5768%28Office.15%29.aspx), [AddConnector](http://msdn.microsoft.com/library/shapes-addconnector-method-powerpoint%28Office.15%29.aspx), [AddCurve](http://msdn.microsoft.com/library/shapes-addcurve-method-powerpoint%28Office.15%29.aspx), [AddLabel](http://msdn.microsoft.com/library/shapes-addlabel-method-powerpoint%28Office.15%29.aspx), [AddLine](http://msdn.microsoft.com/library/shapes-addline-method-powerpoint%28Office.15%29.aspx), [AddMediaObject](http://msdn.microsoft.com/library/shapes-addmediaobject-method-powerpoint%28Office.15%29.aspx), [AddOLEObject](http://msdn.microsoft.com/library/shapes-addoleobject-method-powerpoint%28Office.15%29.aspx), [AddPicture](http://msdn.microsoft.com/library/shapes-addpicture-method-powerpoint%28Office.15%29.aspx), [AddPlaceholder](http://msdn.microsoft.com/library/shapes-addplaceholder-method-powerpoint%28Office.15%29.aspx), [AddPolyline](http://msdn.microsoft.com/library/shapes-addpolyline-method-powerpoint%28Office.15%29.aspx), [AddShape](http://msdn.microsoft.com/library/shapes-addshape-method-powerpoint%28Office.15%29.aspx), [AddTable](http://msdn.microsoft.com/library/shapes-addtable-method-powerpoint%28Office.15%29.aspx), [AddTextbox](http://msdn.microsoft.com/library/shapes-addtextbox-method-powerpoint%28Office.15%29.aspx), [AddTextEffect](http://msdn.microsoft.com/library/shapes-addtexteffect-method-powerpoint%28Office.15%29.aspx), [AddTitle](http://msdn.microsoft.com/library/shapes-addtitle-method-powerpoint%28Office.15%29.aspx).

Use  **Selection.ShapeRange** (index), where index is the shape name or the index number, to return a **Shape** object that represents a shape within the selection. The following example sets the fill for the first shape in the selection in the active window, assuming that there's at least one shape in the selection.




```
ActiveWindow.Selection.ShapeRange(1).Fill _

    .ForeColor.RGB = RGB(255, 0, 0)
```

Use  **Shapes.Title** to return a **Shape** object that represents an existing slide title. Use **Shapes.AddTitle** to add a title to a slide that doesn't already have one and return a **Shape** object that represents the newly created title. Use **Shapes.Placeholders** (index), where index is the placeholder's index number, to return a **Shape** object that represents a placeholder. If you have not changed the layering order of the shapes on a slide, the following three statements are equivalent, assuming that slide one has a title.




```
ActivePresentation.Slides(1).Shapes.Title _

    .TextFrame.TextRange.Font.Italic = True

ActivePresentation.Slides(1).Shapes.Placeholders(1) _

    .TextFrame.TextRange.Font.Italic = True

ActivePresentation.Slides(1).Shapes(1).TextFrame _

    .TextRange.Font.Italic = True
```

To return a  **Shape** object that represents one of the shapes attached by a connector, use the[BeginConnectedShape](http://msdn.microsoft.com/library/connectorformat-beginconnectedshape-property-powerpoint%28Office.15%29.aspx)or [EndConnectedShape](http://msdn.microsoft.com/library/connectorformat-endconnectedshape-property-powerpoint%28Office.15%29.aspx)property.



To return a  **Shape** object that represents the default shape for a presentation, use the[DefaultShape](http://msdn.microsoft.com/library/presentation-defaultshape-property-powerpoint%28Office.15%29.aspx)property.



Use the [BuildFreeform](http://msdn.microsoft.com/library/shapes-buildfreeform-method-powerpoint%28Office.15%29.aspx)and [AddNodes](http://msdn.microsoft.com/library/freeformbuilder-addnodes-method-powerpoint%28Office.15%29.aspx)methods to define the geometry of a new freeform, and use the [ConvertToShape](http://msdn.microsoft.com/library/freeformbuilder-converttoshape-method-powerpoint%28Office.15%29.aspx)method to create the freeform and return the  **Shape** object that represents it.

Use  **GroupItems** (index), where index is the shape name or the index number within the group, to return a **Shape** object that represents a single shape in a grouped shape.

Use the [Group](http://msdn.microsoft.com/library/shaperange-group-method-powerpoint%28Office.15%29.aspx)or [Regroup](http://msdn.microsoft.com/library/shaperange-regroup-method-powerpoint%28Office.15%29.aspx)method to group a range of shapes and return a single  **Shape** object that represents the newly formed group. After a group has been formed, you can work with the group the same way you work with any other shape.


## Methods



|**Name**|
|:-----|
|[Apply](http://msdn.microsoft.com/library/shape-apply-method-powerpoint%28Office.15%29.aspx)|
|[ApplyAnimation](http://msdn.microsoft.com/library/shape-applyanimation-method-powerpoint%28Office.15%29.aspx)|
|[ConvertTextToSmartArt](http://msdn.microsoft.com/library/shape-converttexttosmartart-method-powerpoint%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/shape-copy-method-powerpoint%28Office.15%29.aspx)|
|[Cut](http://msdn.microsoft.com/library/shape-cut-method-powerpoint%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/shape-delete-method-powerpoint%28Office.15%29.aspx)|
|[Duplicate](http://msdn.microsoft.com/library/shape-duplicate-method-powerpoint%28Office.15%29.aspx)|
|[Flip](http://msdn.microsoft.com/library/shape-flip-method-powerpoint%28Office.15%29.aspx)|
|[IncrementLeft](http://msdn.microsoft.com/library/shape-incrementleft-method-powerpoint%28Office.15%29.aspx)|
|[IncrementRotation](http://msdn.microsoft.com/library/shape-incrementrotation-method-powerpoint%28Office.15%29.aspx)|
|[IncrementTop](http://msdn.microsoft.com/library/shape-incrementtop-method-powerpoint%28Office.15%29.aspx)|
|[PickUp](http://msdn.microsoft.com/library/shape-pickup-method-powerpoint%28Office.15%29.aspx)|
|[PickupAnimation](http://msdn.microsoft.com/library/shape-pickupanimation-method-powerpoint%28Office.15%29.aspx)|
|[RerouteConnections](http://msdn.microsoft.com/library/shape-rerouteconnections-method-powerpoint%28Office.15%29.aspx)|
|[ScaleHeight](http://msdn.microsoft.com/library/shape-scaleheight-method-powerpoint%28Office.15%29.aspx)|
|[ScaleWidth](http://msdn.microsoft.com/library/shape-scalewidth-method-powerpoint%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/shape-select-method-powerpoint%28Office.15%29.aspx)|
|[SetShapesDefaultProperties](http://msdn.microsoft.com/library/shape-setshapesdefaultproperties-method-powerpoint%28Office.15%29.aspx)|
|[Ungroup](http://msdn.microsoft.com/library/shape-ungroup-method-powerpoint%28Office.15%29.aspx)|
|[UpgradeMedia](http://msdn.microsoft.com/library/shape-upgrademedia-method-powerpoint%28Office.15%29.aspx)|
|[ZOrder](http://msdn.microsoft.com/library/shape-zorder-method-powerpoint%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[ActionSettings](http://msdn.microsoft.com/library/shape-actionsettings-property-powerpoint%28Office.15%29.aspx)|
|[Adjustments](http://msdn.microsoft.com/library/shape-adjustments-property-powerpoint%28Office.15%29.aspx)|
|[AlternativeText](http://msdn.microsoft.com/library/shape-alternativetext-property-powerpoint%28Office.15%29.aspx)|
|[AnimationSettings](http://msdn.microsoft.com/library/shape-animationsettings-property-powerpoint%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/shape-application-property-powerpoint%28Office.15%29.aspx)|
|[AutoShapeType](http://msdn.microsoft.com/library/shape-autoshapetype-property-powerpoint%28Office.15%29.aspx)|
|[BackgroundStyle](http://msdn.microsoft.com/library/shape-backgroundstyle-property-powerpoint%28Office.15%29.aspx)|
|[BlackWhiteMode](http://msdn.microsoft.com/library/shape-blackwhitemode-property-powerpoint%28Office.15%29.aspx)|
|[Callout](http://msdn.microsoft.com/library/shape-callout-property-powerpoint%28Office.15%29.aspx)|
|[Chart](http://msdn.microsoft.com/library/shape-chart-property-powerpoint%28Office.15%29.aspx)|
|[Child](http://msdn.microsoft.com/library/shape-child-property-powerpoint%28Office.15%29.aspx)|
|[ConnectionSiteCount](http://msdn.microsoft.com/library/shape-connectionsitecount-property-powerpoint%28Office.15%29.aspx)|
|[Connector](http://msdn.microsoft.com/library/shape-connector-property-powerpoint%28Office.15%29.aspx)|
|[ConnectorFormat](http://msdn.microsoft.com/library/shape-connectorformat-property-powerpoint%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/shape-creator-property-powerpoint%28Office.15%29.aspx)|
|[CustomerData](http://msdn.microsoft.com/library/shape-customerdata-property-powerpoint%28Office.15%29.aspx)|
|[Fill](http://msdn.microsoft.com/library/shape-fill-property-powerpoint%28Office.15%29.aspx)|
|[Glow](http://msdn.microsoft.com/library/shape-glow-property-powerpoint%28Office.15%29.aspx)|
|[GroupItems](http://msdn.microsoft.com/library/shape-groupitems-property-powerpoint%28Office.15%29.aspx)|
|[HasChart](http://msdn.microsoft.com/library/shape-haschart-property-powerpoint%28Office.15%29.aspx)|
|[HasInkXML](http://msdn.microsoft.com/library/shape-hasinkxml-property-powerpoint%28Office.15%29.aspx)|
|[HasSmartArt](http://msdn.microsoft.com/library/shape-hassmartart-property-powerpoint%28Office.15%29.aspx)|
|[HasTable](http://msdn.microsoft.com/library/shape-hastable-property-powerpoint%28Office.15%29.aspx)|
|[HasTextFrame](http://msdn.microsoft.com/library/shape-hastextframe-property-powerpoint%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/shape-height-property-powerpoint%28Office.15%29.aspx)|
|[HorizontalFlip](http://msdn.microsoft.com/library/shape-horizontalflip-property-powerpoint%28Office.15%29.aspx)|
|[Id](http://msdn.microsoft.com/library/shape-id-property-powerpoint%28Office.15%29.aspx)|
|[InkXML](http://msdn.microsoft.com/library/shape-inkxml-property-powerpoint%28Office.15%29.aspx)|
|[IsNarration](http://msdn.microsoft.com/library/shape-isnarration-property-powerpoint%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/shape-left-property-powerpoint%28Office.15%29.aspx)|
|[Line](http://msdn.microsoft.com/library/shape-line-property-powerpoint%28Office.15%29.aspx)|
|[LinkFormat](http://msdn.microsoft.com/library/shape-linkformat-property-powerpoint%28Office.15%29.aspx)|
|[LockAspectRatio](http://msdn.microsoft.com/library/shape-lockaspectratio-property-powerpoint%28Office.15%29.aspx)|
|[MediaFormat](http://msdn.microsoft.com/library/shape-mediaformat-property-powerpoint%28Office.15%29.aspx)|
|[MediaType](http://msdn.microsoft.com/library/shape-mediatype-property-powerpoint%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/shape-name-property-powerpoint%28Office.15%29.aspx)|
|[Nodes](http://msdn.microsoft.com/library/shape-nodes-property-powerpoint%28Office.15%29.aspx)|
|[OLEFormat](http://msdn.microsoft.com/library/shape-oleformat-property-powerpoint%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/shape-parent-property-powerpoint%28Office.15%29.aspx)|
|[ParentGroup](http://msdn.microsoft.com/library/shape-parentgroup-property-powerpoint%28Office.15%29.aspx)|
|[PictureFormat](http://msdn.microsoft.com/library/shape-pictureformat-property-powerpoint%28Office.15%29.aspx)|
|[PlaceholderFormat](http://msdn.microsoft.com/library/shape-placeholderformat-property-powerpoint%28Office.15%29.aspx)|
|[Reflection](http://msdn.microsoft.com/library/shape-reflection-property-powerpoint%28Office.15%29.aspx)|
|[Rotation](http://msdn.microsoft.com/library/shape-rotation-property-powerpoint%28Office.15%29.aspx)|
|[Shadow](http://msdn.microsoft.com/library/shape-shadow-property-powerpoint%28Office.15%29.aspx)|
|[ShapeStyle](http://msdn.microsoft.com/library/shape-shapestyle-property-powerpoint%28Office.15%29.aspx)|
|[SmartArt](http://msdn.microsoft.com/library/shape-smartart-property-powerpoint%28Office.15%29.aspx)|
|[SoftEdge](http://msdn.microsoft.com/library/shape-softedge-property-powerpoint%28Office.15%29.aspx)|
|[Table](http://msdn.microsoft.com/library/shape-table-property-powerpoint%28Office.15%29.aspx)|
|[Tags](http://msdn.microsoft.com/library/shape-tags-property-powerpoint%28Office.15%29.aspx)|
|[TextEffect](http://msdn.microsoft.com/library/shape-texteffect-property-powerpoint%28Office.15%29.aspx)|
|[TextFrame](http://msdn.microsoft.com/library/shape-textframe-property-powerpoint%28Office.15%29.aspx)|
|[TextFrame2](http://msdn.microsoft.com/library/shape-textframe2-property-powerpoint%28Office.15%29.aspx)|
|[ThreeD](http://msdn.microsoft.com/library/shape-threed-property-powerpoint%28Office.15%29.aspx)|
|[Title](http://msdn.microsoft.com/library/shape-title-property-powerpoint%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/shape-top-property-powerpoint%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/shape-type-property-powerpoint%28Office.15%29.aspx)|
|[VerticalFlip](http://msdn.microsoft.com/library/shape-verticalflip-property-powerpoint%28Office.15%29.aspx)|
|[Vertices](http://msdn.microsoft.com/library/shape-vertices-property-powerpoint%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/shape-visible-property-powerpoint%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/shape-width-property-powerpoint%28Office.15%29.aspx)|
|[ZOrderPosition](http://msdn.microsoft.com/library/shape-zorderposition-property-powerpoint%28Office.15%29.aspx)|

## See also


<<<<<<< HEAD
#### Other resources

[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/object-model-powerpoint-vba-reference%28Office.15%29.aspx)
=======
#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)
>>>>>>> d7667e83d23dbf8ebf5bf068ba6fed14c840c0f5

