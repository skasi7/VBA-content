---
title: Series Members (PowerPoint)
ms.prod: POWERPOINT
ms.assetid: f7e7168d-3c6f-20db-1e75-56a101c69a70
---


# Series Members (PowerPoint)
Represents a series in a chart.

Represents a series in a chart.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[ApplyDataLabels](series-applydatalabels-method-powerpoint.md)|Applies data labels to a series.|
|[ClearFormats](series-clearformats-method-powerpoint.md)|Clears the formatting of the object.|
|[Copy](series-copy-method-powerpoint.md)|If the series has a picture fill, copies the picture to the Clipboard.|
|[DataLabels](series-datalabels-method-powerpoint.md)|Returns an object that represents either a single data label (a  **[DataLabel](datalabel-object-powerpoint.md)** object) or a collection of all the data labels for the series (a **[DataLabels](datalabels-object-powerpoint.md)** collection).|
|[Delete](series-delete-method-powerpoint.md)|Deletes the object.|
|[ErrorBar](series-errorbar-method-powerpoint.md)|Applies error bars to the series. |
|[Paste](series-paste-method-powerpoint.md)|Pastes a picture from the Clipboard as the marker on the selected series.|
|[Points](series-points-method-powerpoint.md)|Returns a collection of all the points in the series.|
|[Select](series-select-method-powerpoint.md)|Selects the object.|
|[Trendlines](series-trendlines-method-powerpoint.md)|Returns a collection of all the trendlines for the series.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](series-application-property-powerpoint.md)|When used without an object qualifier, returns an  **[Application](application-object-powerpoint.md)** object that represents the Microsoft PowerPoint application. When used with an object qualifier, returns an **Application** object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.|
|[ApplyPictToEnd](series-applypicttoend-property-powerpoint.md)|**True** if a picture is applied to the end of the point or all points in the series. Read/write **Boolean**.|
|[ApplyPictToFront](series-applypicttofront-property-powerpoint.md)|**True** if a picture is applied to the front of the point or all points in the series. Read/write **Boolean**.|
|[ApplyPictToSides](series-applypicttosides-property-powerpoint.md)|**True** if a picture is applied to the sides of the point or all points in the series. Read/write **Boolean**.|
|[AxisGroup](series-axisgroup-property-powerpoint.md)|Returns the type of axis group. Read/write  **[XlAxisGroup](xlaxisgroup-enumeration-powerpoint.md)**.|
|[BarShape](series-barshape-property-powerpoint.md)|Returns or sets the shape used for a single series in a 3-D bar or column chart. Read/write  **[XlBarShape](xlbarshape-enumeration-powerpoint.md)**.|
|[BubbleSizes](series-bubblesizes-property-powerpoint.md)|Returns or sets a string that refers to the worksheet cells that contain the x-value, y-value, and size data for the bubble chart. Read/write  **Variant**.|
|[ChartType](series-charttype-property-powerpoint.md)|Returns or sets the chart type. Read/write  **[XlChartType](xlcharttype-enumeration-excel.md)**.|
|[Creator](series-creator-property-powerpoint.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.|
|[ErrorBars](series-errorbars-property-powerpoint.md)|Returns the error bars for the series. Read-only  **[ErrorBars](errorbars-object-powerpoint.md)**.|
|[Explosion](series-explosion-property-powerpoint.md)|Returns or sets the explosion value for a pie-chart or doughnut-chart slice. Read/write  **Long**.|
|[Format](series-format-property-powerpoint.md)|Returns the line, fill, and effect formatting for the object. Read-only  **[ChartFormat](chartformat-object-powerpoint.md)**.|
|[Formula](series-formula-property-powerpoint.md)|Returns or sets the object's formula in A1-style notation and in the language of the macro. Read/write  **String**.|
|[FormulaLocal](series-formulalocal-property-powerpoint.md)|Returns or sets the formula for the object, using A1-style references in the language of the user. Read/write  **String**.|
|[FormulaR1C1](series-formular1c1-property-powerpoint.md)|Returns or sets the formula for the object, using R1C1-style notation in the language of the macro. Read/write  **String**.|
|[FormulaR1C1Local](series-formular1c1local-property-powerpoint.md)|Returns or sets the formula for the object, using R1C1-style notation in the language of the user. Read/write  **String**.|
|[Has3DEffect](series-has3deffect-property-powerpoint.md)|**True** if the series has a three-dimensional appearance. Read/write **Boolean**.|
|[HasDataLabels](series-hasdatalabels-property-powerpoint.md)|**True** if the series has data labels. Read/write **Boolean**.|
|[HasErrorBars](series-haserrorbars-property-powerpoint.md)|**True** if the series has error bars. Read/write **Boolean**.|
|[HasLeaderLines](series-hasleaderlines-property-powerpoint.md)|**True** if the series has leader lines. Read/write **Boolean**.|
|[InvertColor](series-invertcolor-property-powerpoint.md)|Returns or sets the fill color for negative data points in a series. Read/write.|
|[InvertColorIndex](series-invertcolorindex-property-powerpoint.md)|Returns or sets the fill color for negative data points in a series. Read/write.|
|[InvertIfNegative](series-invertifnegative-property-powerpoint.md)|**True** if Microsoft Word inverts the pattern in the object when it corresponds to a negative number. Read/write **Variant**.|
|[IsFiltered](series-isfiltered-property-powerpoint.md)|Returns or sets a  **Boolean** that determines whether the specified chart series is filtered out from the chart. Read-write.|
|[LeaderLines](series-leaderlines-property-powerpoint.md)|Returns the leader lines for the series. Read-only  **[LeaderLines](leaderlines-object-powerpoint.md)**.|
|[MarkerBackgroundColor](series-markerbackgroundcolor-property-powerpoint.md)|Sets the marker background color as an RGB value or returns the corresponding color index value. Read/write  **Long**.|
|[MarkerBackgroundColorIndex](series-markerbackgroundcolorindex-property-powerpoint.md)|Returns or sets the marker background color as an index into the current color palette, or as one of the following  **[XlColorIndex](xlcolorindex-enumeration-powerpoint.md)** constants: **xlColorIndexAutomatic** or **xlColorIndexNone**. Read/write **Long**.|
|[MarkerForegroundColor](series-markerforegroundcolor-property-powerpoint.md)|Sets the marker foreground color as an RGB value or returns the corresponding color index value. Read/write  **Long**.|
|[MarkerForegroundColorIndex](series-markerforegroundcolorindex-property-powerpoint.md)|Returns or sets the marker foreground color as an index into the current color palette, or as one of the following  **[XlColorIndex](xlcolorindex-enumeration-powerpoint.md)** constants: **xlColorIndexAutomatic** or **xlColorIndexNone**. Read/write **Long**.|
|[MarkerSize](series-markersize-property-powerpoint.md)|Returns or sets the data-marker size, in points. Read/write  **Long**.|
|[MarkerStyle](series-markerstyle-property-powerpoint.md)|Returns or sets the marker style for a point or series in a line chart, scatter chart, or radar chart. Read/write  **[XlMarkerStyle](xlmarkerstyle-enumeration-powerpoint.md)**.|
|[Name](series-name-property-powerpoint.md)|Returns or sets the name of the object. Read/write  **String**.|
|[Parent](series-parent-property-powerpoint.md)|Returns the parent for the specified object. Read-only  **Object**.|
|**[ParentDataLabelOption](series-parentdatalabeloption-property-powerpoint.md)**|
|:-----|
|[PictureType](series-picturetype-property-powerpoint.md)|Returns or sets a value that specifies how pictures are displayed on a column or bar picture chart. Read/write  **[XlChartPictureType](xlchartpicturetype-enumeration-powerpoint.md)**.|
|[PictureUnit2](series-pictureunit2-property-powerpoint.md)|Returns or sets the unit for each picture on the chart if the  **[PictureType](series-picturetype-property-powerpoint.md)** property is set to **xlStackScale**; otherwise, this property is ignored. Read/write **Double**.|
|[PlotColorIndex](series-plotcolorindex-property-powerpoint.md)|Returns an index value that is used internally to associate series formatting with chart elements. Read-only.|
|[PlotOrder](series-plotorder-property-powerpoint.md)|Returns or sets the plot order for the selected series within the chart group. Read/write  **Long**.|
|[QuartileCalculationInclusiveMedian](series-quartilecalculationinclusivemedian-property-powerpoint.md)|**True** if the series uses an inclusive median quartile calculation method. Read/write **Boolean**.|
|[Shadow](series-shadow-property-powerpoint.md)|Returns or sets a value that indicates whether the object has a shadow. Read/write  **Boolean**.|
|[Smooth](series-smooth-property-powerpoint.md)|**True** if curve smoothing is enabled for the line chart or scatter chart. Read/write **Boolean**.|
|[Type](series-type-property-powerpoint.md)|Returns or sets the series type. Read/write  **Long**.|
|[Values](series-values-property-powerpoint.md)|Returns or sets a collection of all the values in the series. Read/write  **Variant**.|
|[XValues](series-xvalues-property-powerpoint.md)|Returns or sets an array of x values for a chart series. Read/write  **Variant**.|

