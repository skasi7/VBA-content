---
title: CalloutFormat Members (PowerPoint)
ms.prod: POWERPOINT
ms.assetid: 2c1284aa-3540-a0b2-15cd-ef6c87fd8b67
---


# CalloutFormat Members (PowerPoint)
Contains properties and methods that apply to line callouts.

Contains properties and methods that apply to line callouts.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[AutomaticLength](calloutformat-automaticlength-method-powerpoint.md)|Specifies that the first segment of the callout line (the segment attached to the text callout box) be scaled automatically when the callout is moved. Use the  **[CustomLength](calloutformat-customlength-method-powerpoint.md)** method to specify that the first segment of the callout line retain the fixed length returned by the **Length** property whenever the callout is moved. Applies only to callouts whose lines consist of more than one segment (types **msoCalloutThree** and **msoCalloutFour** ).|
|[CustomDrop](calloutformat-customdrop-method-powerpoint.md)|Sets the vertical distance (in points) from the edge of the text bounding box to the place where the callout line attaches to the text box. This distance is measured from the top of the text box unless the  **AutoAttach** property is set to **True** and the text box is to the left of the origin of the callout line (the place that the callout points to). In this case the drop distance is measured from the bottom of the text box.|
|[CustomLength](calloutformat-customlength-method-powerpoint.md)|Specifies that the first segment of the callout line (the segment attached to the text callout box) retain a fixed length whenever the callout is moved. |
|[PresetDrop](calloutformat-presetdrop-method-powerpoint.md)|Specifies whether the callout line attaches to the top, bottom, or center of the callout text box or whether it attaches at a point that's a specified distance from the top or bottom of the text box.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Accent](calloutformat-accent-property-powerpoint.md)|Determines whether a vertical accent bar separates the callout text from the callout line. Read/write.|
|[Angle](calloutformat-angle-property-powerpoint.md)|Returns or sets the angle of the callout line. If the callout line contains more than one line segment, this property returns or sets the angle of the segment that is farthest from the callout text box. Read/write.|
|[Application](calloutformat-application-property-powerpoint.md)|Returns an  **[Application](application-object-powerpoint.md)** object that represents the creator of the specified object.|
|[AutoAttach](calloutformat-autoattach-property-powerpoint.md)|Determines whether the place where the callout line attaches to the callout text box changes, depending on whether the origin of the callout line (where the callout points to) is to the left or right of the callout text box. Read/write.|
|[AutoLength](calloutformat-autolength-property-powerpoint.md)|Determines whether the first segment of the callout retains the fixed length specified by the  **[Length](calloutformat-length-property-powerpoint.md)** property, or is scaled automatically, whenever the callout is moved. Read-only.|
|[Border](calloutformat-border-property-powerpoint.md)|Determines whether the text in the specified callout is surrounded by a border. Read/write.|
|[Creator](calloutformat-creator-property-powerpoint.md)|Returns a  **Long** that represents the four-character creator code for the application in which the specified object was created. For example, if the object was created in Microsoft PowerPoint, this property returns the hexadecimal number 50575054. Read-only.|
|[Drop](calloutformat-drop-property-powerpoint.md)|For callouts with an explicitly set drop value, this property returns the vertical distance (in points) from the edge of the text bounding box to the place where the callout line attaches to the text box. Read-only.|
|[DropType](calloutformat-droptype-property-powerpoint.md)|Returns a value that indicates where the callout line attaches to the callout text box. Read-only.|
|[Gap](calloutformat-gap-property-powerpoint.md)|Returns or sets the horizontal distance (in points) between the end of the callout line and the text bounding box. Read/write.|
|[Length](calloutformat-length-property-powerpoint.md)|When the  **[AutoLength](calloutformat-autolength-property-powerpoint.md)** property of the specified callout is set to **False**, the **Length** property returns the length (in points) of the first segment of the callout line (the segment attached to the text callout box). Read-only.|
|[Parent](calloutformat-parent-property-powerpoint.md)|Returns the parent object for the specified object.|
|[Type](calloutformat-type-property-powerpoint.md)|Represents the type of callout. Read/write.|

