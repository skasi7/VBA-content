---
title: CalloutFormat Members (Excel)
ms.prod: EXCEL
ms.assetid: 29203369-3128-3336-6e78-d1853c4619a6
---


# CalloutFormat Members (Excel)
Contains properties and methods that apply to line callouts.

Contains properties and methods that apply to line callouts.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[AutomaticLength](calloutformat-automaticlength-method-excel.md)|Specifies that the first segment of the callout line (the segment attached to the text callout box) be scaled automatically when the callout is moved. Use the  **[CustomLength](calloutformat-customlength-method-excel.md)** method to specify that the first segment of the callout line retain the fixed length returned by the **[Length](calloutformat-length-property-excel.md)** property whenever the callout is moved. Applies only to callouts whose lines consist of more than one segment (types **msoCalloutThree** and **msoCalloutFour** ).|
|[CustomDrop](calloutformat-customdrop-method-excel.md)|Sets the vertical distance (in points) from the edge of the text bounding box to the place where the callout line attaches to the text box. This distance is measured from the top of the text box unless the  **[AutoAttach](calloutformat-autoattach-property-excel.md)** property is set to **True** and the text box is to the left of the origin of the callout line (the place that the callout points to), in which case the drop distance is measured from the bottom of the text box.|
|[CustomLength](calloutformat-customlength-method-excel.md)|Specifies that the first segment of the callout line (the segment attached to the text callout box) retain a fixed length whenever the callout is moved. Use the  **[AutomaticLength](calloutformat-automaticlength-method-excel.md)** method to specify that the first segment of the callout line be scaled automatically whenever the callout is moved. Applies only to callouts whose lines consist of more than one segment (types **msoCalloutThree** and **msoCalloutFour** ).|
|[PresetDrop](calloutformat-presetdrop-method-excel.md)|Specifies whether the callout line attaches to the top, bottom, or center of the callout text box or whether it attaches at a point that's a specified distance from the top or bottom of the text box.|

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

