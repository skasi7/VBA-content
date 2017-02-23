---
title: LegendKey Properties (Word)
ms.prod: WORD
ms.assetid: 67be6bf1-7727-4ea4-8d48-2525dab8e4b5
---


# LegendKey Properties (Word)

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](legendkey-application-property-word.md)|When used without an object qualifier, returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application. When used with an object qualifier, returns an **Application** object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.|
|[Creator](legendkey-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[Format](legendkey-format-property-word.md)|Returns the line, fill, and effect formatting for the object. Read-only  **[ChartFormat](chartformat-object-word.md)** .|
|[Height](legendkey-height-property-word.md)|Returns the height, in points, of the object. Read-only  **Double** .|
|[InvertIfNegative](legendkey-invertifnegative-property-word.md)| **True** if Microsoft Word inverts the pattern in the object when it corresponds to a negative number. Read/write **Variant** .|
|[Left](legendkey-left-property-word.md)|Returns the distance, in points, from the left edge of the object to the left edge of the chart area. Read-only  **Double** .|
|[MarkerBackgroundColor](legendkey-markerbackgroundcolor-property-word.md)|Sets the marker background color as an RGB value or returns the corresponding color index value. Read/write  **Long** .|
|[MarkerBackgroundColorIndex](legendkey-markerbackgroundcolorindex-property-word.md)|Returns or sets the marker background color as an index into the current color palette, or as one of the following  **[XlColorIndex](xlcolorindex-enumeration-word.md)** constants: **xlColorIndexAutomatic** or **xlColorIndexNone** . Read/write **Long** .|
|[MarkerForegroundColor](legendkey-markerforegroundcolor-property-word.md)|Sets the marker foreground color as an RGB value or returns the corresponding color index value. Read/write  **Long** .|
|[MarkerForegroundColorIndex](legendkey-markerforegroundcolorindex-property-word.md)|Returns or sets the marker foreground color as an index into the current color palette, or as one of the following  **[XlColorIndex](xlcolorindex-enumeration-word.md)** constants: **xlColorIndexAutomatic** or **xlColorIndexNone** . Read/write **Long** .|
|[MarkerSize](legendkey-markersize-property-word.md)|Returns or sets the data-marker size, in points. Read/write  **Long** .|
|[MarkerStyle](legendkey-markerstyle-property-word.md)|Returns or sets the marker style for a point or series in a line chart, scatter chart, or radar chart. Read/write  **[XlMarkerStyle](xlmarkerstyle-enumeration-word.md)** .|
|[Parent](legendkey-parent-property-word.md)|Returns the parent for the specified object. Read-only  **Object** .|
|[PictureType](legendkey-picturetype-property-word.md)|Returns or sets the way pictures are displayed on a legend key. Read/write  **[XlChartPictureType](xlchartpicturetype-enumeration-word.md)** .|
|[PictureUnit2](legendkey-pictureunit2-property-word.md)|Returns or sets the unit for each picture on the chart if the  **[PictureType](legendkey-picturetype-property-word.md)** property is set to **xlStackScale** ; otherwise, this property is ignored. Read/write **Double** .|
|[Shadow](legendkey-shadow-property-word.md)|Returns or sets a value that indicates whether the object has a shadow. Read/write  **Boolean** .|
|[Smooth](legendkey-smooth-property-word.md)| **True** if curve smoothing is turned on for the legend key. Read/write **Boolean** .|
|[Top](legendkey-top-property-word.md)|Returns the distance, in points, from the top edge of the object to the top of the first row (on a worksheet) or the top of the chart area (on a chart). Read-only  **Double** .|
|[Width](legendkey-width-property-word.md)|Returns the width, in points, of the object. Read-only  **Double** .|

