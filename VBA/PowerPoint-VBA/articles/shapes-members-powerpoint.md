---
title: Shapes Members (PowerPoint)
ms.prod: POWERPOINT
ms.assetid: 75a4880e-71e1-fe10-a719-f7c13389a74e
---


# Shapes Members (PowerPoint)
A collection of all the  **[Shape](shape-object-powerpoint.md)** objects on the specified slide.

A collection of all the  **[Shape](shape-object-powerpoint.md)** objects on the specified slide.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[AddCallout](shapes-addcallout-method-powerpoint.md)|Creates a borderless line callout. Returns a  **[Shape](shape-object-powerpoint.md)** object that represents the new callout.|
|[AddChart2](shapes-addchart2-method-powerpoint.md)|Adds a chart to the document. Returns a [Shape](shape-object-powerpoint.md) object that represents a chart and adds it to the specified collection.|
|[AddConnector](shapes-addconnector-method-powerpoint.md)|Creates a connector. Returns a  **[Shape](shape-object-powerpoint.md)** object that represents the new connector. When a connector is added, it is not connected to anything. Use the **[BeginConnect](connectorformat-beginconnect-method-powerpoint.md)** and **[EndConnect](connectorformat-endconnect-method-powerpoint.md)** methods to attach the beginning and end of a connector to other shapes in the document.|
|[AddCurve](shapes-addcurve-method-powerpoint.md)|Creates a B?zier curve. Returns a  **[Shape](shape-object-powerpoint.md)** object that represents the new curve.|
|[AddInkShapeFromXML](shapes-addinkshapefromxml-method-powerpoint.md)|Creates an ink shape. Returns a [Shape](shape-object-powerpoint.md) object that represents the new ink shape.|
|[AddLabel](shapes-addlabel-method-powerpoint.md)|Creates a label. Returns a  **[Shape](shape-object-powerpoint.md)** object that represents the new label.|
|[AddLine](shapes-addline-method-powerpoint.md)|Creates a line. Returns a  **[Shape](shape-object-powerpoint.md)** object that represents the new line.|
|[AddMediaObject2](shapes-addmediaobject2-method-powerpoint.md)|Replaces deprecated [Shapes.AddMediaObject Method (PowerPoint)](shapes-addmediaobject-method-powerpoint.md). Adds a new media object. |
|[AddMediaObjectFromEmbedTag](shapes-addmediaobjectfromembedtag-method-powerpoint.md)|Adds a media object from an embedded tag to a  **Shapes** object.|
|[AddOLEObject](shapes-addoleobject-method-powerpoint.md)|Creates an OLE object. Returns a  **[Shape](shape-object-powerpoint.md)** object that represents the new OLE object.|
|[AddPicture](shapes-addpicture-method-powerpoint.md)|Creates a picture from an existing file. Returns a  **[Shape](shape-object-powerpoint.md)** object that represents the new picture.|
|[AddPicture2](shapes-addpicture2-method-powerpoint.md)|Creates a picture from an existing file. Returns a  **[Shape](shape-object-powerpoint.md)** object that represents the new picture.|
|[AddPlaceholder](shapes-addplaceholder-method-powerpoint.md)|Restores a previously deleted placeholder on a slide. Returns a  **[Shape](shape-object-powerpoint.md)** object that represents the restored placeholder.|
|[AddPolyline](shapes-addpolyline-method-powerpoint.md)|Creates an open polyline or a closed polygon drawing. Returns a  **[Shape](shape-object-powerpoint.md)** object that represents the new polyline or polygon.|
|[AddShape](shapes-addshape-method-powerpoint.md)|Creates an AutoShape. Returns a  **[Shape](shape-object-powerpoint.md)** object that represents the new AutoShape.|
|[AddSmartArt](shapes-addsmartart-method-powerpoint.md)|Adds a SmartArt diagram to the  **Shapes** object.|
|[AddTable](shapes-addtable-method-powerpoint.md)|Adds a table shape to a slide.|
|[AddTextbox](shapes-addtextbox-method-powerpoint.md)|Creates a text box. Returns a  **[Shape](shape-object-powerpoint.md)** object that represents the new text box.|
|[AddTextEffect](shapes-addtexteffect-method-powerpoint.md)|Creates a WordArt object. Returns a  **[Shape](shape-object-powerpoint.md)** object that represents the new WordArt object.|
|[AddTitle](shapes-addtitle-method-powerpoint.md)|Restores a previously deleted title placeholder to a slide. Returns a  **[Shape](shape-object-powerpoint.md)** object that represents the restored title.|
|[BuildFreeform](shapes-buildfreeform-method-powerpoint.md)|Builds a freeform object. Returns a  **[FreeformBuilder](freeformbuilder-object-powerpoint.md)** object that represents the freeform as it is being built.|
|[Item](shapes-item-method-powerpoint.md)|Returns a single  **Shape** object from the specified **Shapes** collection.|
|[Paste](shapes-paste-method-powerpoint.md)|Pastes the shapes, slides, or text on the Clipboard into the specified  **Shapes** collection, at the top of the z-order. Each pasted object becomes a member of the specified **Shapes** collection. If the Clipboard contains entire slides, the slides will be pasted as shapes that contain the images of the slides. If the Clipboard contains a text range, the text will be pasted into a newly created **TextFrame** shape. Returns a **[ShapeRange](shaperange-object-powerpoint.md)** object that represents the pasted objects.|
|[PasteSpecial](shapes-pastespecial-method-powerpoint.md)|Pastes the contents of the Clipboard, using a special format.|
|[Range](shapes-range-method-powerpoint.md)|Returns a  **[ShapeRange](shaperange-object-powerpoint.md)** object that represents a subset of the shapes in a **[Shapes](shapes-object-powerpoint.md)** collection.|
|[SelectAll](shapes-selectall-method-powerpoint.md)|Selects all the shapes in a  **[Shapes](shapes-object-powerpoint.md)** collection.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](shapes-application-property-powerpoint.md)|Returns an  **[Application](application-object-powerpoint.md)** object that represents the creator of the specified object.|
|[Count](shapes-count-property-powerpoint.md)|Returns the number of objects in the specified collection. Read-only.|
|[Creator](shapes-creator-property-powerpoint.md)|Returns a  **Long** that represents the four-character creator code for the application in which the specified object was created. For example, if the object was created in Microsoft PowerPoint, this property returns the hexadecimal number 50575054. Read-only.|
|[HasTitle](shapes-hastitle-property-powerpoint.md)|Returns whether the collection of objects on the specified slide contains a title placeholder. Read-only.|
|[Parent](shapes-parent-property-powerpoint.md)|Returns the parent object for the specified object.|
|[Placeholders](shapes-placeholders-property-powerpoint.md)|Returns a  **[Placeholders](placeholders-object-powerpoint.md)** collection that represents the collection of all the placeholders on a slide. Read-only.|
|[Title](shapes-title-property-powerpoint.md)|Returns a  **[Shape](shape-object-powerpoint.md)** object that represents the slide title. Read-only.|

