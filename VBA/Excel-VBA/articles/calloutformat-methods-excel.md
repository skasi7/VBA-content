---
title: CalloutFormat Methods (Excel)
ms.prod: EXCEL
ms.assetid: 89796632-2bbf-40ad-9e2b-a33ce2d5d75a
---


# CalloutFormat Methods (Excel)

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[AutomaticLength](calloutformat-automaticlength-method-excel.md)|Specifies that the first segment of the callout line (the segment attached to the text callout box) be scaled automatically when the callout is moved. Use the  **[CustomLength](calloutformat-customlength-method-excel.md)** method to specify that the first segment of the callout line retain the fixed length returned by the **[Length](calloutformat-length-property-excel.md)** property whenever the callout is moved. Applies only to callouts whose lines consist of more than one segment (types **msoCalloutThree** and **msoCalloutFour** ).|
|[CustomDrop](calloutformat-customdrop-method-excel.md)|Sets the vertical distance (in points) from the edge of the text bounding box to the place where the callout line attaches to the text box. This distance is measured from the top of the text box unless the  **[AutoAttach](calloutformat-autoattach-property-excel.md)** property is set to **True** and the text box is to the left of the origin of the callout line (the place that the callout points to), in which case the drop distance is measured from the bottom of the text box.|
|[CustomLength](calloutformat-customlength-method-excel.md)|Specifies that the first segment of the callout line (the segment attached to the text callout box) retain a fixed length whenever the callout is moved. Use the  **[AutomaticLength](calloutformat-automaticlength-method-excel.md)** method to specify that the first segment of the callout line be scaled automatically whenever the callout is moved. Applies only to callouts whose lines consist of more than one segment (types **msoCalloutThree** and **msoCalloutFour** ).|
|[PresetDrop](calloutformat-presetdrop-method-excel.md)|Specifies whether the callout line attaches to the top, bottom, or center of the callout text box or whether it attaches at a point that's a specified distance from the top or bottom of the text box.|

