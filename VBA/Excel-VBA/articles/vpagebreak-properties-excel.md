---
title: VPageBreak Properties (Excel)
ms.prod: EXCEL
ms.assetid: 421ede28-ac79-4852-81af-caa81c0e39a5
---


# VPageBreak Properties (Excel)

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](vpagebreak-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[Creator](vpagebreak-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[Extent](vpagebreak-extent-property-excel.md)|Returns the type of the specified page break: full-screen or only within a print area. Can be either of the following  **[XlPageBreakExtent](xlpagebreakextent-enumeration-excel.md)** constants: **xlPageBreakFull** or **xlPageBreakPartial** . Read-only **Long** .|
|[Location](vpagebreak-location-property-excel.md)|Returns or sets the cell (a  **Range** object) that defines the page-break location. Horizontal page breaks are aligned with the top edge of the location cell; vertical page breaks are aligned with the left edge of the location cell. Read/write **[Range](range-object-excel.md)** .|
|[Parent](vpagebreak-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[Type](vpagebreak-type-property-excel.md)|Returns or sets a  **[XlPageBreak](xlpagebreak-enumeration-excel.md)** value that represents the page break type.|

