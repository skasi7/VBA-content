---
title: TextFrame Members (Excel)
ms.prod: EXCEL
ms.assetid: 299ac22a-bf3d-11ca-90e8-a05d52a760d4
---


# TextFrame Members (Excel)
Represents the text frame in a  **[Shape](shape-object-excel.md)** object. Contains the text in the text frame as well as the properties and methods that control the alignment and anchoring of the text frame.

Represents the text frame in a  **[Shape](shape-object-excel.md)** object. Contains the text in the text frame as well as the properties and methods that control the alignment and anchoring of the text frame.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Characters](textframe-characters-method-excel.md)|Returns a  **[Characters](characters-object-excel.md)** object that represents a range of characters within a shape's text frame. You can use the **Characters** object to add and format characters within the text frame.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](textframe-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[AutoMargins](textframe-automargins-property-excel.md)|Returns or sets whether Excel automatically calculates text frame margins. Read/write|
|[AutoSize](textframe-autosize-property-excel.md)| **True** if the size of the specified object is changed automatically to fit text within its boundaries. Read/write **Boolean** .|
|[Creator](textframe-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[HorizontalAlignment](textframe-horizontalalignment-property-excel.md)|Returns or sets a  **[XlHAlign](xlhalign-enumeration-excel.md)** value that represents the horizontal alignment for the specified object.|
|[HorizontalOverflow](textframe-horizontaloverflow-property-excel.md)|Returns or sets the horizontal overflow setting for the specified object. Read/write|
|[MarginBottom](textframe-marginbottom-property-excel.md)|Returns or sets the distance (in points) between the bottom of the text frame and the bottom of the inscribed rectangle of the shape that contains the text. Read/write  **Single** .|
|[MarginLeft](textframe-marginleft-property-excel.md)|Returns or sets the distance (in points) between the left edge of the text frame and the left edge of the inscribed rectangle of the shape that contains the text. Read/write  **Single** .|
|[MarginRight](textframe-marginright-property-excel.md)|Returns or sets the distance (in points) between the right edge of the text frame and the right edge of the inscribed rectangle of the shape that contains the text. Read/write  **Single** .|
|[MarginTop](textframe-margintop-property-excel.md)|Returns or sets the distance (in points) between the top of the text frame and the top of the inscribed rectangle of the shape that contains the text. Read/write  **Single** .|
|[Orientation](textframe-orientation-property-excel.md)|Returns or sets a  **Long** value that represents the text frame orientation.|
|[Parent](textframe-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[ReadingOrder](textframe-readingorder-property-excel.md)|Returns or sets the reading order for the specified object. Can be one of the following constants:  **xlRTL** (right-to-left), **xlLTR** (left-to-right), or **xlContext** . Read/write **Long** .|
|[VerticalAlignment](textframe-verticalalignment-property-excel.md)|Returns or sets a  **[XlVAlign](xlvalign-enumeration-excel.md)** value that represents the vertical alignment of the specified object.|
|[VerticalOverflow](textframe-verticaloverflow-property-excel.md)|Returns or sets the vertical overflow setting for the specified object. Read/write|

