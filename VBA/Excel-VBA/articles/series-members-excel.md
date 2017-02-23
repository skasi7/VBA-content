---
title: Series Members (Excel)
ms.prod: EXCEL
ms.assetid: eeab4f69-b436-9de7-5d4a-0a5c63f2dfce
---


# Series Members (Excel)
Represents a series in a chart.

Represents a series in a chart.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[ApplyDataLabels](series-applydatalabels-method-excel.md)|Applies data labels to a series.|
|[ClearFormats](series-clearformats-method-excel.md)|Clears the formatting of the object.|
|[Copy](series-copy-method-excel.md)|If the series has a picture fill, then this method copies the picture to the Clipboard.|
|[DataLabels](series-datalabels-method-excel.md)|Returns an object that represents either a single data label (a  **[DataLabel](datalabel-object-excel.md)** object) or a collection of all the data labels for the series (a **[DataLabels](datalabels-object-excel.md)** collection).|
|[Delete](series-delete-method-excel.md)|Deletes the object.|
|[ErrorBar](series-errorbar-method-excel.md)|Applies error bars to the series.  **Variant** .|
|[Paste](series-paste-method-excel.md)|Pastes a picture from the Clipboard as the marker on the selected series.|
|[Points](series-points-method-excel.md)|Returns an object that represents a single point (a  **[Point](point-object-excel.md)** object) or a collection of all the points (a **[Points](points-object-excel.md)** collection) in the series. Read-only|
|[Select](series-select-method-excel.md)|Selects the object.|
|[Trendlines](series-trendlines-method-excel.md)|Returns an object that represents a single trendline (a  **[Trendline](trendline-object-excel.md)** object) or a collection of all the trendlines (a **[Trendlines](trendlines-object-excel.md)** collection) for the series.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](series-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[ApplyPictToEnd](series-applypicttoend-property-excel.md)| **True** if a picture is applied to the end of the point or all points in the series. Read/write **Boolean** .|
|[ApplyPictToFront](series-applypicttofront-property-excel.md)| **True** if a picture is applied to the front of the point or all points in the series. Read/write **Boolean** .|
|[ApplyPictToSides](series-applypicttosides-property-excel.md)| **True** if a picture is applied to the sides of the point or all points in the series. Read/write **Boolean** .|
|[AxisGroup](series-axisgroup-property-excel.md)|Returns or sets the group for the specified series. Read/write|
|[BarShape](series-barshape-property-excel.md)|Returns or sets the shape used with the 3-D bar or column chart. Read/write  **[XlBarShape](xlbarshape-enumeration-excel.md)** .|
|[BubbleSizes](series-bubblesizes-property-excel.md)|Returns or sets a string that refers to the worksheet cells containing the x-value, y-value and size data for the bubble chart. When you return the cell reference, it will return a string describing the cells in A1-style notation. To set the size data for the bubble chart, you must use R1C1-style notation. Applies only to bubble charts. Read/write  **Variant** .|
|[ChartType](series-charttype-property-excel.md)|Returns or sets the chart type. Read/write  **[XlChartType](xlcharttype-enumeration-excel.md)** .|
|[Creator](series-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[ErrorBars](series-errorbars-property-excel.md)|Returns an  **[ErrorBars](errorbars-object-excel.md)** object that represents the error bars for the series. Read-only.|
|[Explosion](series-explosion-property-excel.md)|Returns or sets the explosion value for a pie-chart or doughnut-chart slice. Returns 0 (zero) if there's no explosion (the tip of the slice is in the center of the pie). Read/write  **Long** .|
|[Format](series-format-property-excel.md)|Returns the  **[ChartFormat](chartformat-object-excel.md)** object. Read-only.|
|[Formula](series-formula-property-excel.md)|Returns or sets a  **String** value that represents the object's formula in A1-style notation and in the language of the macro.|
|[FormulaLocal](series-formulalocal-property-excel.md)|Returns or sets the formula for the object, using A1-style references in the language of the user. Read/write  **String** .|
|[FormulaR1C1](series-formular1c1-property-excel.md)|Returns or sets the formula for the object, using R1C1-style notation in the language of the macro. Read/write  **String** .|
|[FormulaR1C1Local](series-formular1c1local-property-excel.md)|Returns or sets the formula for the object, using R1C1-style notation in the language of the user. Read/write  **String** .|
|[Has3DEffect](series-has3deffect-property-excel.md)| **True** if the series has a three-dimensional appearance. Read/write **Boolean** .|
|[HasDataLabels](series-hasdatalabels-property-excel.md)| **True** if the series has data labels. Read/write **Boolean** .|
|[HasErrorBars](series-haserrorbars-property-excel.md)| **True** if the series has error bars. This property isn't available for 3-D charts. Read/write **Boolean** .|
|[HasLeaderLines](series-hasleaderlines-property-excel.md)| **True** if the series has leader lines. Read/write **Boolean** .|
|[InvertColor](series-invertcolor-property-excel.md)|Returns or sets the fill color for negative data points in a series. Read/write|
|[InvertColorIndex](series-invertcolorindex-property-excel.md)|Returns or sets the fill color for negative data points in a series. Read/write|
|[InvertIfNegative](series-invertifnegative-property-excel.md)| **True** if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number. Read/write **Boolean** .|
|[IsFiltered](series-isfiltered-property-excel.md)|This setting controls whether the series has been filtered out from the chart. The default value is  **False** . **Boolean** Read/Write.|
|[LeaderLines](series-leaderlines-property-excel.md)|Returns a  **LeaderLines** object that represents the leader lines for the series. Read-only.|
|[MarkerBackgroundColor](series-markerbackgroundcolor-property-excel.md)|Sets the marker background color as an RGB value or returns the corresponding color index value. Applies only to line, scatter, and radar charts. Read/write  **Long** .|
|[MarkerBackgroundColorIndex](series-markerbackgroundcolorindex-property-excel.md)|Returns or sets the marker background color as an index into the current color palette, or as one of the following  **[XlColorIndex](xlcolorindex-enumeration-excel.md)** constants: **xlColorIndexAutomatic** or **xlColorIndexNone** . Applies only to line, scatter, and radar charts. Read/write **Long** .|
|[MarkerForegroundColor](series-markerforegroundcolor-property-excel.md)|Sets the marker foreground color as an RGB value or returns the corresponding color index value. Applies only to line, scatter, and radar charts. Read/write  **Long** .|
|[MarkerForegroundColorIndex](series-markerforegroundcolorindex-property-excel.md)|Returns or sets the marker foreground color as an index into the current color palette, or as one of the following  **[XlColorIndex](xlcolorindex-enumeration-excel.md)** constants: **xlColorIndexAutomatic** or **xlColorIndexNone** . Applies only to line, scatter, and radar charts. Read/write **Long** .|
|[MarkerSize](series-markersize-property-excel.md)|Returns or sets the data-marker size, in points. Can be a value from 2 through 72. Read/write  **Long** .|
|[MarkerStyle](series-markerstyle-property-excel.md)|Returns or sets the marker style for a point or series in a line chart, scatter chart, or radar chart. Read/write  **[XlMarkerStyle](xlmarkerstyle-enumeration-excel.md)** .|
|[Name](series-name-property-excel.md)|Returns or sets a  **String** value representing the name of the object.|
|[Parent](series-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[PictureType](series-picturetype-property-excel.md)|Returns or sets a  **[XlChartPictureType](xlchartpicturetype-enumeration-excel.md)** value that represents the way pictures are displayed on a column or bar picture chart.|
|[PictureUnit2](series-pictureunit2-property-excel.md)|Returns or sets the unit for each picture on the chart if the  **[PictureType](series-picturetype-property-excel.md)** property is set to **xlStackScale** (if not, this property is ignored). Read/write **Double** .|
|[PlotColorIndex](series-plotcolorindex-property-excel.md)|Returns an index value that is used internally to associate series formatting with chart elements. Read-only|
|[PlotOrder](series-plotorder-property-excel.md)|Returns or sets the plot order for the selected series within the chart group. Read/write  **Long** .|
|[Shadow](series-shadow-property-excel.md)|Returns or sets a  **Boolean** value that determines if the object has a shadow.|
|[Smooth](series-smooth-property-excel.md)| **True** if curve smoothing is turned on for the line chart or scatter chart. Applies only to line and scatter charts. Read/write.|
|[Type](series-type-property-excel.md)|Returns or sets a Long value that represents the series type.|
|[Values](series-values-property-excel.md)|Returns or sets a  **Variant** value that represents a collection of all the values in the series.|
|[XValues](series-xvalues-property-excel.md)|Returns or sets an array of x values for a chart series. The  **XValues** property can be set to a range on a worksheet or to an array of values, but it cannot be a combination of both. Read/write **Variant** .|
|[ParentDataLabelOption](series-parentdatalabeloption-property-excel.md)|Specifies the parent data label option (banner, overlapping, or none) for the specified series within the chart group. Read/write [XLParentDataLabelOptions](xlparentdatalabeloptions-enumeration-excel.md).|
|[QuartileCalculationInclusiveMedian](series-quartilecalculationinclusivemedian-property-excel.md)| **True** if the series uses an inclusive median quartile calculation method. Read/write **Boolean**.|

