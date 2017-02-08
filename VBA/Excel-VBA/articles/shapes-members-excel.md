---
title: Shapes Members (Excel)
ms.prod: EXCEL
ms.assetid: f5d0be42-46cc-2916-8953-401e50a5cef7
---


# Shapes Members (Excel)
A collection of all the  **[Shape](shape-object-excel.md)** objects on the specified sheet.

A collection of all the  **[Shape](shape-object-excel.md)** objects on the specified sheet.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[AddCallout](shapes-addcallout-method-excel.md)| Creates a borderless line callout. Returns a **[Shape](shape-object-excel.md)** object that represents the new callout.|
|[AddChart2](shapes-addchart2-method-excel.md)|Adds a chart to the document. Returns a  **Shape** object that represents a chart and adds it to the specified collection.|
|[AddConnector](shapes-addconnector-method-excel.md)|Creates a connector. Returns a  **[Shape](shape-object-excel.md)** object that represents the new connector. When a connector is added, it's not connected to anything. Use the **[BeginConnect](connectorformat-beginconnect-method-excel.md)** and **[EndConnect](connectorformat-endconnect-method-excel.md)** methods to attach the beginning and end of a connector to other shapes in the document.|
|[AddCurve](shapes-addcurve-method-excel.md)|Returns a  **[Shape](shape-object-excel.md)** object that represents a Bézier curve in a worksheet.|
|[AddFormControl](shapes-addformcontrol-method-excel.md)|Creates a Microsoft Excel control. Returns a  **[Shape](shape-object-excel.md)** object that represents the new control.|
|[AddLabel](shapes-addlabel-method-excel.md)|Creates a label. Returns a  **[Shape](shape-object-excel.md)** object that represents the new label.|
|[AddLine](shapes-addline-method-excel.md)|As it applies to the  **Shapes** object, returns a **[Shape](shape-object-excel.md)** object that represents the new line in a worksheet.|
|[AddOLEObject](shapes-addoleobject-method-excel.md)|Creates an OLE object. Returns a  **[Shape](shape-object-excel.md)** object that represents the new OLE object.|
|[AddPicture](shapes-addpicture-method-excel.md)|Creates a picture from an existing file. Returns a  **Shape** object that represents the new picture.|
|[AddPicture2](shapes-addpicture2-method-excel.md)||
|[AddPolyline](shapes-addpolyline-method-excel.md)|Creates an open polyline or a closed polygon drawing. Returns a  **[Shape](shape-object-excel.md)** object that represents the new polyline or polygon.|
|[AddShape](shapes-addshape-method-excel.md)|Returns a  **[Shape](shape-object-excel.md)** object that represents the new AutoShape in a worksheet.|
|[AddSmartArt](shapes-addsmartart-method-excel.md)|Creates a new SmartArt graphic with the specified layout. |
|[AddTextbox](shapes-addtextbox-method-excel.md)|Creates a text box. Returns a  **[Shape](shape-object-excel.md)** object that represents the new text box.|
|[AddTextEffect](shapes-addtexteffect-method-excel.md)|Creates a WordArt object. Returns a  **[Shape](shape-object-excel.md)** object that represents the new WordArt object.|
|[BuildFreeform](shapes-buildfreeform-method-excel.md)|Builds a freeform object. Returns a  **[FreeformBuilder](freeformbuilder-object-excel.md)** object that represents the freeform as it is being built. Use the **[AddNodes](freeformbuilder-addnodes-method-excel.md)** method to add segments to the freeform. After you have added at least one segment to the freeform, you can use the **[ConvertToShape](freeformbuilder-converttoshape-method-excel.md)** method to convert the **FreeformBuilder** object into a **[Shape](shape-object-excel.md)** object that has the geometric description you've defined in the **FreeformBuilder** object.|
|[Item](shapes-item-method-excel.md)|Returns a single object from a collection.|
|[SelectAll](shapes-selectall-method-excel.md)|Selects all the shapes in the specified  **[Shapes](shapes-object-excel.md)** collection.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](shapes-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[Count](shapes-count-property-excel.md)|Returns a  **Long** value that represents the number of objects in the collection.|
|[Creator](shapes-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[Parent](shapes-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[Range](shapes-range-property-excel.md)|Returns a  **[ShapeRange](shaperange-object-excel.md)** object that represents a subset of the shapes in a **Shapes** collection.|

