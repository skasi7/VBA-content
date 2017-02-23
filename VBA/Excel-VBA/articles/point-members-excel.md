---
title: Point Members (Excel)
ms.prod: EXCEL
ms.assetid: a533258d-fc3b-9fe1-2a77-a55ecbe7bd7a
---


# Point Members (Excel)
Represents a single point in a series in a chart.

Represents a single point in a series in a chart.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[ApplyDataLabels](point-applydatalabels-method-excel.md)|Applies data labels to a point.|
|[ClearFormats](point-clearformats-method-excel.md)|Clears the formatting of the object.|
|[Copy](point-copy-method-excel.md)|If the point has a picture fill, then this method copies the picture to the Clipboard.|
|[Delete](point-delete-method-excel.md)|Deletes the series the point belongs to.|
|[Paste](point-paste-method-excel.md)|Pastes a picture from the Clipboard as the marker on the selected point.|
|[PieSliceLocation](point-pieslicelocation-method-excel.md)|Returns the vertical or horizontal position of a point on a chart item, in points, from the top or left edge of the object to the top or left edge of the chart area.|
|[Select](point-select-method-excel.md)|Selects the object.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](point-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[ApplyPictToEnd](point-applypicttoend-property-excel.md)| **True** if a picture is applied to the end of the point or all points in the series. Read/write **Boolean** .|
|[ApplyPictToFront](point-applypicttofront-property-excel.md)| **True** if a picture is applied to the front of the point or all points in the series. Read/write **Boolean** .|
|[ApplyPictToSides](point-applypicttosides-property-excel.md)| **True** if a picture is applied to the sides of the point or all points in the series. Read/write **Boolean** .|
|[Creator](point-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[DataLabel](point-datalabel-property-excel.md)|Returns a  **[DataLabel](datalabel-object-excel.md)** object that represents the data label associated with the point. Read-only.|
|[Explosion](point-explosion-property-excel.md)|Returns or sets the explosion value for a pie-chart or doughnut-chart slice. Returns 0 (zero) if there's no explosion (the tip of the slice is in the center of the pie). Read/write  **Long** .|
|[Format](point-format-property-excel.md)|Returns the  **[ChartFormat](chartformat-object-excel.md)** object. Read-only.|
|[Has3DEffect](point-has3deffect-property-excel.md)| **True** if a point has a three-dimensional appearance. Read/write **Boolean** .|
|[HasDataLabel](point-hasdatalabel-property-excel.md)| **True** if the point has a data label. Read/write **Boolean** .|
|[Height](point-height-property-excel.md)|Returns the height, in points, of the object. Read-only.|
|[InvertIfNegative](point-invertifnegative-property-excel.md)| **True** if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number. Read/write **Boolean** .|
|[Left](point-left-property-excel.md)|Returns a value that represents the distance, in points, from the left edge of the object to the left edge of the chart area. Read-only.|
|[MarkerBackgroundColor](point-markerbackgroundcolor-property-excel.md)|Sets the marker background color as an RGB value or returns the corresponding color index value. Applies only to line, scatter, and radar charts. Read/write  **Long** .|
|[MarkerBackgroundColorIndex](point-markerbackgroundcolorindex-property-excel.md)|Returns or sets the marker background color as an index into the current color palette, or as one of the following  **[XlColorIndex](xlcolorindex-enumeration-excel.md)** constants: **xlColorIndexAutomatic** or **xlColorIndexNone** . Applies only to line, scatter, and radar charts. Read/write **Long** .|
|[MarkerForegroundColor](point-markerforegroundcolor-property-excel.md)|Sets the marker foreground color as an RGB value or returns the corresponding color index value. Applies only to line, scatter, and radar charts. Read/write  **Long** .|
|[MarkerForegroundColorIndex](point-markerforegroundcolorindex-property-excel.md)|Returns or sets the marker foreground color as an index into the current color palette, or as one of the following  **[XlColorIndex](xlcolorindex-enumeration-excel.md)** constants: **xlColorIndexAutomatic** or **xlColorIndexNone** . Applies only to line, scatter, and radar charts. Read/write **Long** .|
|[MarkerSize](point-markersize-property-excel.md)|Returns or sets the data-marker size, in points. Can be a value from 2 through 72. Read/write  **Long** .|
|[MarkerStyle](point-markerstyle-property-excel.md)|Returns or sets the marker style for a point or series in a line chart, scatter chart, or radar chart. Read/write  **[XlMarkerStyle](xlmarkerstyle-enumeration-excel.md)** .|
|[Name](point-name-property-excel.md)|Returns the object name. Read-only.|
|[Parent](point-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[PictureType](point-picturetype-property-excel.md)|Returns or sets a  **[XlChartPictureType](xlchartpicturetype-enumeration-excel.md)** value that represents the way pictures are displayed on a column or bar picture chart.|
|[PictureUnit2](point-pictureunit2-property-excel.md)|Returns or sets the unit for each picture on the chart if the  **[PictureType](point-picturetype-property-excel.md)** property is set to **xlStackScale** (if not, this property is ignored). Read/write **Double** .|
|[SecondaryPlot](point-secondaryplot-property-excel.md)| **True** if the point is in the secondary section of either a pie of pie chart or a bar of pie chart. Applies only to points on pie of pie charts or bar of pie charts. Read/write **Boolean** .|
|[Shadow](point-shadow-property-excel.md)|Returns or sets a  **Boolean** value that determines if the object has a shadow.|
|[Top](point-top-property-excel.md)|Returns a value that represents the distance, in points, from the top edge of the object to the top edge of the chart area. Read-only.|
|[Width](point-width-property-excel.md)|Returns the width, in points, of the object. Read-only.|
|[IsTotal](point-istotal-property-excel.md)| **True** if the point represents a total. Read/write **Boolean**.|

