---
title: Shapes Object (PowerPoint)
keywords: vbapp10.chm543000
f1_keywords:
- vbapp10.chm543000
ms.prod: POWERPOINT
ms.assetid: eb208855-254e-1a0f-884b-4a5edcfd584d
---


# Shapes Object (PowerPoint)

A collection of all the  **[Shape](http://msdn.microsoft.com/library/shape-object-powerpoint%28Office.15%29.aspx)** objects on the specified slide.


## Remarks

Each  **Shape** object represents an object in the drawing layer, such as an AutoShape, freeform, OLE object, or picture.


 **Note**  If you want to work with a subset of the shapes on a document — for example, to do something to only the AutoShapes on the document or to only the selected shapes — you must construct a  **[ShapeRange](shaperange-object-powerpoint.md)** collection that contains the shapes you want to work with. For an overview of how to work either with a single shape or with more than one shape at a time, see[How to: Work with Shapes (Drawing Objects)](http://msdn.microsoft.com/library/work-with-shapes-drawing-objects%28Office.15%29.aspx).


## Example

Use the  **Shapes** property to return the **Shapes** collection. The following example selects all the shapes in the active presentation.


```
ActivePresentation.Slides(1).Shapes.SelectAll
```


 **Note**  If you want to do something (like delete or set a property) to all the shapes on a document at the same time, use the [Range](http://msdn.microsoft.com/library/shapes-range-method-powerpoint%28Office.15%29.aspx)method with no argument to create a  **ShapeRange** object that contains all the shapes in the **Shapes** collection, and then apply the appropriate property or method to the **ShapeRange** object.

Use the [AddCallout](http://msdn.microsoft.com/library/shapes-addcallout-method-powerpoint%28Office.15%29.aspx), [AddComment](http://msdn.microsoft.com/library/11347ca1-cef3-0923-2544-cb80e7fc5768%28Office.15%29.aspx), [AddConnector](http://msdn.microsoft.com/library/shapes-addconnector-method-powerpoint%28Office.15%29.aspx), [AddCurve](http://msdn.microsoft.com/library/shapes-addcurve-method-powerpoint%28Office.15%29.aspx), [AddLabel](http://msdn.microsoft.com/library/shapes-addlabel-method-powerpoint%28Office.15%29.aspx), [AddLine](http://msdn.microsoft.com/library/shapes-addline-method-powerpoint%28Office.15%29.aspx), [AddMediaObject](http://msdn.microsoft.com/library/shapes-addmediaobject-method-powerpoint%28Office.15%29.aspx), [AddOLEObject](http://msdn.microsoft.com/library/shapes-addoleobject-method-powerpoint%28Office.15%29.aspx), [AddPicture](http://msdn.microsoft.com/library/shapes-addpicture-method-powerpoint%28Office.15%29.aspx), [AddPlaceholder](http://msdn.microsoft.com/library/shapes-addplaceholder-method-powerpoint%28Office.15%29.aspx), [AddPolyline](http://msdn.microsoft.com/library/shapes-addpolyline-method-powerpoint%28Office.15%29.aspx), [AddShape](http://msdn.microsoft.com/library/shapes-addshape-method-powerpoint%28Office.15%29.aspx), [AddTable](http://msdn.microsoft.com/library/shapes-addtable-method-powerpoint%28Office.15%29.aspx), [AddTextbox](http://msdn.microsoft.com/library/shapes-addtextbox-method-powerpoint%28Office.15%29.aspx), [AddTextEffect](http://msdn.microsoft.com/library/shapes-addtexteffect-method-powerpoint%28Office.15%29.aspx), or [AddTitle](http://msdn.microsoft.com/library/shapes-addtitle-method-powerpoint%28Office.15%29.aspx)method to create a new shape and add it to the  **Shapes** collection. Use the[BuildFreeform](http://msdn.microsoft.com/library/shapes-buildfreeform-method-powerpoint%28Office.15%29.aspx)method in conjunction with the [ConvertToShape](http://msdn.microsoft.com/library/freeformbuilder-converttoshape-method-powerpoint%28Office.15%29.aspx)method to create a new freeform and add it to the collection. The following example adds a rectangle to the active presentation.




```
ActivePresentation.Slides(1).Shapes.AddShape Type:=msoShapeRectangle, _

    Left:=50, Top:=50, Width:=100, Height:=200
```

Use  **Shapes** (index), where index is the shape's name or index number, to return a single **Shape** object. The following example sets the fill to a preset shade for shape one in the active presentation.




```
ActivePresentation.Slides(1).Shapes(1).Fill _

    .PresetGradient Style:=msoGradientHorizontal, Variant:=1, _

    PresetGradientType:=msoGradientBrass
```

Use  **Shapes.Range** (index), where index is the shape's name or index number or an array of shape names or index numbers, to return a **[ShapeRange](shaperange-object-powerpoint.md)** collection that represents a subset of the **Shapes** collection. The following example sets the fill pattern for shapes one and three in the active presentation.




```
ActivePresentation.Slides(1).Shapes.Range(Array(1, 3)).Fill _

    .Patterned Pattern:=msoPatternHorizontalBrick
```

Use  **Shapes.Placeholders** (index), where index is the placeholder number, to return a **Shape** object that represents a placeholder. If the specified slide has a title, use **Shapes.Placeholders(1)** or **Shapes.Title** to return the title placeholder. The following example adds a slide to the active presentation and then adds text to both the title and the subtitle (the subtitle is the second placeholder on a slide with this layout).




```
With ActivePresentation.Slides.Add(Index:=1, Layout:=ppLayoutTitle).Shapes

    .Title.TextFrame.TextRange = "This is the title text"

    .Placeholders(2).TextFrame.TextRange = "This is subtitle text"

End With
```


## Methods



|**Name**|
|:-----|
|[AddCallout](http://msdn.microsoft.com/library/shapes-addcallout-method-powerpoint%28Office.15%29.aspx)|
|[AddChart2](http://msdn.microsoft.com/library/shapes-addchart2-method-powerpoint%28Office.15%29.aspx)|
|[AddConnector](http://msdn.microsoft.com/library/shapes-addconnector-method-powerpoint%28Office.15%29.aspx)|
|[AddCurve](http://msdn.microsoft.com/library/shapes-addcurve-method-powerpoint%28Office.15%29.aspx)|
|[AddInkShapeFromXML](http://msdn.microsoft.com/library/shapes-addinkshapefromxml-method-powerpoint%28Office.15%29.aspx)|
|[AddLabel](http://msdn.microsoft.com/library/shapes-addlabel-method-powerpoint%28Office.15%29.aspx)|
|[AddLine](http://msdn.microsoft.com/library/shapes-addline-method-powerpoint%28Office.15%29.aspx)|
|[AddMediaObject2](http://msdn.microsoft.com/library/shapes-addmediaobject2-method-powerpoint%28Office.15%29.aspx)|
|[AddMediaObjectFromEmbedTag](http://msdn.microsoft.com/library/shapes-addmediaobjectfromembedtag-method-powerpoint%28Office.15%29.aspx)|
|[AddOLEObject](http://msdn.microsoft.com/library/shapes-addoleobject-method-powerpoint%28Office.15%29.aspx)|
|[AddPicture](http://msdn.microsoft.com/library/shapes-addpicture-method-powerpoint%28Office.15%29.aspx)|
|[AddPicture2](http://msdn.microsoft.com/library/shapes-addpicture2-method-powerpoint%28Office.15%29.aspx)|
|[AddPlaceholder](http://msdn.microsoft.com/library/shapes-addplaceholder-method-powerpoint%28Office.15%29.aspx)|
|[AddPolyline](http://msdn.microsoft.com/library/shapes-addpolyline-method-powerpoint%28Office.15%29.aspx)|
|[AddShape](http://msdn.microsoft.com/library/shapes-addshape-method-powerpoint%28Office.15%29.aspx)|
|[AddSmartArt](http://msdn.microsoft.com/library/shapes-addsmartart-method-powerpoint%28Office.15%29.aspx)|
|[AddTable](http://msdn.microsoft.com/library/shapes-addtable-method-powerpoint%28Office.15%29.aspx)|
|[AddTextbox](http://msdn.microsoft.com/library/shapes-addtextbox-method-powerpoint%28Office.15%29.aspx)|
|[AddTextEffect](http://msdn.microsoft.com/library/shapes-addtexteffect-method-powerpoint%28Office.15%29.aspx)|
|[AddTitle](http://msdn.microsoft.com/library/shapes-addtitle-method-powerpoint%28Office.15%29.aspx)|
|[BuildFreeform](http://msdn.microsoft.com/library/shapes-buildfreeform-method-powerpoint%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/shapes-item-method-powerpoint%28Office.15%29.aspx)|
|[Paste](http://msdn.microsoft.com/library/shapes-paste-method-powerpoint%28Office.15%29.aspx)|
|[PasteSpecial](http://msdn.microsoft.com/library/shapes-pastespecial-method-powerpoint%28Office.15%29.aspx)|
|[Range](http://msdn.microsoft.com/library/shapes-range-method-powerpoint%28Office.15%29.aspx)|
|[SelectAll](http://msdn.microsoft.com/library/shapes-selectall-method-powerpoint%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/shapes-application-property-powerpoint%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/shapes-count-property-powerpoint%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/shapes-creator-property-powerpoint%28Office.15%29.aspx)|
|[HasTitle](http://msdn.microsoft.com/library/shapes-hastitle-property-powerpoint%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/shapes-parent-property-powerpoint%28Office.15%29.aspx)|
|[Placeholders](http://msdn.microsoft.com/library/shapes-placeholders-property-powerpoint%28Office.15%29.aspx)|
|[Title](http://msdn.microsoft.com/library/shapes-title-property-powerpoint%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/object-model-powerpoint-vba-reference%28Office.15%29.aspx)
