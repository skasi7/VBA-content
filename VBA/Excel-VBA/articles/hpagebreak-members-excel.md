---
title: HPageBreak Members (Excel)
ms.prod: EXCEL
ms.assetid: 32b561ff-a0cf-142b-0a46-c622a42b6125
---


# HPageBreak Members (Excel)
Represents a horizontal page break. 

Represents a horizontal page break. 


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Delete](hpagebreak-delete-method-excel.md)|Deletes the object.|
|[DragOff](hpagebreak-dragoff-method-excel.md)|Drags a page break out of the print area.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](hpagebreak-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[Creator](hpagebreak-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[Extent](hpagebreak-extent-property-excel.md)|Returns the type of the specified page break: full-screen or only within a print area. Can be either of the following  **[XlPageBreakExtent](xlpagebreakextent-enumeration-excel.md)** constants: **xlPageBreakFull** or **xlPageBreakPartial** . Read-only **Long** .|
|[Location](hpagebreak-location-property-excel.md)|Returns or sets the cell (a  **Range** object) that defines the page-break location. Horizontal page breaks are aligned with the top edge of the location cell; vertical page breaks are aligned with the left edge of the location cell. Read/write **[Range](range-object-excel.md)** .|
|[Parent](hpagebreak-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[Type](hpagebreak-type-property-excel.md)|Returns or sets a  **[XlPageBreak](xlpagebreak-enumeration-excel.md)** value that represents the page break type.|

