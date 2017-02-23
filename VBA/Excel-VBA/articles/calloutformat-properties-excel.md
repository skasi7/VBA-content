---
title: CalloutFormat Properties (Excel)
ms.prod: EXCEL
ms.assetid: 329d65b1-e1b5-4a12-a022-46d9bdfbc672
---


# CalloutFormat Properties (Excel)

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Accent](calloutformat-accent-property-excel.md)|Allows the user to place a vertical accent bar to separate the callout text from the callout line. Read/write  **[MsoTriState](msotristate-enumeration-office.md)** .|
|[Angle](calloutformat-angle-property-excel.md)|Returns or sets the angle of the callout line. If the callout line contains more than one line segment, this property returns or sets the angle of the segment that is farthest from the callout text box. Read/write  **[MsoCalloutAngleType](msocalloutangletype-enumeration-office.md)** .|
|[Application](calloutformat-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[AutoAttach](calloutformat-autoattach-property-excel.md)| **True** if the place where the callout line attaches to the callout text box changes depending on whether the origin of the callout line (where the callout points to) is to the left or right of the callout text box. Read/write **[MsoTriState](msotristate-enumeration-office.md)** .|
|[AutoLength](calloutformat-autolength-property-excel.md)|Applies only to callouts whose lines consist of more than one segment (types  **msoCalloutThree** and **msoCalloutFour** ). Read/write **[MsoTriState](msotristate-enumeration-office.md)** .|
|[Border](calloutformat-border-property-excel.md)|Returns or sets a  **[MsoTriState](msotristate-enumeration-office.md)** value that represents the visibility options for the border of the object.|
|[Creator](calloutformat-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[Drop](calloutformat-drop-property-excel.md)|For callouts with an explicitly set drop value, this property returns the vertical distance (in points) from the edge of the text bounding box to the place where the callout line attaches to the text box. Read-only  **Single** .|
|[DropType](calloutformat-droptype-property-excel.md)|Returns a value that indicates where the callout line attaches to the callout text box. Read-only  **[MsoCalloutDropType](msocalloutdroptype-enumeration-office.md)** .|
|[Gap](calloutformat-gap-property-excel.md)|Returns or sets the horizontal distance (in points) between the end of the callout line and the text bounding box. Read/write  **Single** .|
|[Length](calloutformat-length-property-excel.md)|Returns a  **Single** value that represents the length (in points) of the first segment of the callout line (the segment attached to the text callout box.)|
|[Parent](calloutformat-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[Type](calloutformat-type-property-excel.md)|Returns or sets a  **[MsoCalloutType](msocallouttype-enumeration-office.md)** value that represents the callout format type.|

