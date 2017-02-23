---
title: OLEObjects Members (Excel)
ms.prod: EXCEL
ms.assetid: 7c3b0619-a988-1b8c-51b1-4c8ef3180264
---


# OLEObjects Members (Excel)
A collection of all the  **[OLEObject](oleobject-object-excel.md)** objects on the specified worksheet.

A collection of all the  **[OLEObject](oleobject-object-excel.md)** objects on the specified worksheet.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Add](oleobjects-add-method-excel.md)|Adds a new OLE object to a sheet. |
|[BringToFront](oleobjects-bringtofront-method-excel.md)|Brings the object to the front of the z-order.|
|[Copy](oleobjects-copy-method-excel.md)|Copies the object to the Clipboard.|
|[CopyPicture](oleobjects-copypicture-method-excel.md)|Copies the selected object to the Clipboard as a picture.  **Variant** .|
|[Cut](oleobjects-cut-method-excel.md)|Cuts the object to the Clipboard.|
|[Delete](oleobjects-delete-method-excel.md)|Deletes the object.|
|[Duplicate](oleobjects-duplicate-method-excel.md)|Duplicates the object and returns a reference to the new copy.|
|[Item](oleobjects-item-method-excel.md)|Returns a single object from a collection.|
|[Select](oleobjects-select-method-excel.md)|Selects the object.|
|[SendToBack](oleobjects-sendtoback-method-excel.md)|Sends the object to the back of the z-order.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](oleobjects-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[AutoLoad](oleobjects-autoload-property-excel.md)| **True** if the OLE object is automatically loaded when the workbook that contains it is opened. Read/write **Boolean** .|
|[Border](oleobjects-border-property-excel.md)|Returns a  **[Border](border-object-excel.md)** object that represents the border of the object.|
|[Count](oleobjects-count-property-excel.md)|Returns a  **Long** value that represents the number of objects in the collection.|
|[Creator](oleobjects-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[Enabled](oleobjects-enabled-property-excel.md)| **True** if the object is enabled. Read/write **Boolean** .|
|[Height](oleobjects-height-property-excel.md)|Returns or sets a  **Double** value that represents the height, in points, of the object.|
|[Interior](oleobjects-interior-property-excel.md)|Returns an  **[Interior](interior-object-excel.md)** object that represents the interior of the specified object.|
|[Left](oleobjects-left-property-excel.md)|Returns or sets a  **Double** value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).|
|[Locked](oleobjects-locked-property-excel.md)|Returns or sets a  **Boolean** value that indicates if the object is locked.|
|[Parent](oleobjects-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[Placement](oleobjects-placement-property-excel.md)|Returns or sets a  **Variant** value, containing an **[XlPlacement](xlplacement-enumeration-excel.md)** constant, that represents the way the object is attached to the cells below it.|
|[PrintObject](oleobjects-printobject-property-excel.md)| **True** if the object will be printed when the document is printed. Read/write **Boolean** .|
|[Shadow](oleobjects-shadow-property-excel.md)|Returns or sets a  **Boolean** value that determines if the object has a shadow.|
|[ShapeRange](oleobjects-shaperange-property-excel.md)|Returns a  **[ShapeRange](shaperange-object-excel.md)** object that represents the specified object or objects. Read-only.|
|[SourceName](oleobjects-sourcename-property-excel.md)|Returns or sets a  **String** value that represents the specified object's link source name.|
|[Top](oleobjects-top-property-excel.md)|Returns or sets a  **Double** value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).|
|[Visible](oleobjects-visible-property-excel.md)|Returns or sets a  **Boolean** value that determines whether the object is visible. Read/write.|
|[Width](oleobjects-width-property-excel.md)|Returns or sets a  **Double** value that represents the width, in points, of the object.|
|[ZOrder](oleobjects-zorder-property-excel.md)|Returns the z-order position of the object. Read-only  **Long** .|

