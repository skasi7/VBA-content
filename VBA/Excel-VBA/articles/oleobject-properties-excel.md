---
title: OLEObject Properties (Excel)
ms.prod: EXCEL
ms.assetid: fa360d25-5127-488d-84a5-1f82c455f3c3
---


# OLEObject Properties (Excel)

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](oleobject-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[AutoLoad](oleobject-autoload-property-excel.md)| **True** if the OLE object is automatically loaded when the workbook that contains it is opened. Read/write **Boolean** .|
|[AutoUpdate](oleobject-autoupdate-property-excel.md)| **True** if the OLE object is updated automatically when the source changes. Valid only if the object is linked (its **OLEType** property must be **xlOLELink** ). Read-only **Boolean** .|
|[Border](oleobject-border-property-excel.md)|Returns a  **[Border](border-object-excel.md)** object that represents the border of the object.|
|[BottomRightCell](oleobject-bottomrightcell-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents the cell that lies under the lower-right corner of the object. Read-only.|
|[Creator](oleobject-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[Enabled](oleobject-enabled-property-excel.md)| **True** if the object is enabled. Read/write **Boolean** .|
|[Height](oleobject-height-property-excel.md)|Returns or sets a  **Double** value that represents the height, in points, of the object.|
|[Index](oleobject-index-property-excel.md)|Returns a  **Long** value that represents the index number of the object within the collection of similar objects.|
|[Interior](oleobject-interior-property-excel.md)|Returns an  **[Interior](interior-object-excel.md)** object that represents the interior of the specified object.|
|[Left](oleobject-left-property-excel.md)|Returns or sets a  **Double** value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).|
|[LinkedCell](oleobject-linkedcell-property-excel.md)|Returns or sets the worksheet range linked to the control's value. If you place a value in the cell, the control takes this value. Likewise, if you change the value of the control, that value is also placed in the cell. Read/write  **String** .|
|[ListFillRange](oleobject-listfillrange-property-excel.md)|Returns or sets the worksheet range used to fill the specified list box. Setting this property destroys any existing list in the list box. Read/write  **String** .|
|[Locked](oleobject-locked-property-excel.md)|Returns or sets a  **Boolean** value that indicates if the object is locked.|
|[Name](oleobject-name-property-excel.md)|Returns or sets a  **String** value representing the name of the object.|
|[Object](oleobject-object-property-excel.md)|Returns the OLE Automation object associated with this OLE object. Read-only  **Object** .|
|[OLEType](oleobject-oletype-property-excel.md)|Returns the OLE object type. Can be one of the following  **XlOLEType** constants: **xlOLELink** or **xlOLEEmbed** . Returns **xlOLELink** if the object is linked (it exists outside of the file), or returns **xlOLEEmbed** if the object is embedded (it's entirely contained within the file). Read-only **Long** .|
|[Parent](oleobject-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[Placement](oleobject-placement-property-excel.md)|Returns or sets a  **Variant** value, containing an **[XlPlacement](xlplacement-enumeration-excel.md)** constant, that represents the way the object is attached to the cells below it.|
|[PrintObject](oleobject-printobject-property-excel.md)| **True** if the object will be printed when the document is printed. Read/write **Boolean** .|
|[progID](oleobject-progid-property-excel.md)|Returns the programmatic identifiers for the object. Read-only  **String** .|
|[Shadow](oleobject-shadow-property-excel.md)|Returns or sets a  **Boolean** value that determines if the object has a shadow.|
|[ShapeRange](oleobject-shaperange-property-excel.md)|Returns a  **[ShapeRange](shaperange-object-excel.md)** object that represents the specified object or objects. Read-only.|
|[SourceName](oleobject-sourcename-property-excel.md)|Returns or sets a  **String** value that represents the specified object's link source name.|
|[Top](oleobject-top-property-excel.md)|Returns or sets a  **Double** value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).|
|[TopLeftCell](oleobject-topleftcell-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents the cell that lies under the upper-left corner of the specified object. Read-only.|
|[Visible](oleobject-visible-property-excel.md)|Returns or sets a  **Boolean** value that determines whether the object is visible. Read/write.|
|[Width](oleobject-width-property-excel.md)|Returns or sets a  **Double** value that represents the width, in points, of the object.|
|[ZOrder](oleobject-zorder-property-excel.md)|Returns the z-order position of the object. Read-only  **Long** .|

