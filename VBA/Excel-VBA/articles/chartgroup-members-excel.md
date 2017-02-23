---
title: ChartGroup Members (Excel)
ms.prod: EXCEL
ms.assetid: 2d31f7af-d639-c8f4-0714-08fc618ec92d
---


# ChartGroup Members (Excel)
Represents one or more series plotted in a chart with the same format.

Represents one or more series plotted in a chart with the same format.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[CategoryCollection](chartgroup-categorycollection-method-excel.md)|Returns an object that represents a collection of all of the visible categories (a [CategoryCollection](categorycollection-object-excel.md) collection) in the chart group.|
|[FullCategoryCollection](chartgroup-fullcategorycollection-method-excel.md)|Returns an object that represents a collection of all of the visible and filtered categories (a [CategoryCollection](categorycollection-object-excel.md) collection) in the chart group.|
|[SeriesCollection](chartgroup-seriescollection-method-excel.md)|Returns an object that represents either a single series (a  **[Series](series-object-excel.md)** object) or a collection of all the series (a **[SeriesCollection](seriescollection-object-excel.md)** collection) in the chart or chart group.|
|||

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](chartgroup-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[AxisGroup](chartgroup-axisgroup-property-excel.md)|Returns or sets the group for the specified chart. Read/write|
|[BinsCountValue](chartgroup-binscountvalue-property-excel.md)|Specifies the number of bins in the histogram chart. Read/write  **Long**.|
|[BinsOverflowEnabled](chartgroup-binsoverflowenabled-property-excel.md)|Specifies whether a bin for values above the [BinsOverflowValue](chartgroup-binsoverflowvalue-property-excel.md) is enabled. Read/write **Boolean**.|
|[BinsOverflowValue](chartgroup-binsoverflowvalue-property-excel.md)|If an [BinsOverflowEnabled](chartgroup-binsoverflowenabled-property-excel.md) is **True**, specifies the value above which an overflow bin is displayed. Read/write  **Double**.|
|[BinsType](chartgroup-binstype-property-excel.md)|Specifies how the horizontal axis of the histogram chart is formatted, by bins type. Read/write [XlBinsType](xlbinstype-enumeration-excel.md).|
|[BinsUnderflowEnabled](chartgroup-binsunderflowenabled-property-excel.md)|Specifies whether a bin for values below the [BinsUnderflowValue](chartgroup-binsunderflowvalue-property-excel.md) is enabled. Read/write **Boolean**.|
|[BinsUnderflowValue](chartgroup-binsunderflowvalue-property-excel.md)|If an [BinsUnderflowEnabled](chartgroup-binsunderflowenabled-property-excel.md) is **True**, specifies the value below which an underflow bin is displayed. Read/write  **Double**.|
|[BinWidthValue](chartgroup-binwidthvalue-property-excel.md)|Specifies the number of points in each range. Read/write  **Double**.|
|[BubbleScale](chartgroup-bubblescale-property-excel.md)|Returns or sets the scale factor for bubbles in the specified chart group. Can be an integer value from 0 (zero) to 300, corresponding to a percentage of the default size. Applies only to bubble charts. Read/write  **Long** .|
|[Creator](chartgroup-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[DoughnutHoleSize](chartgroup-doughnutholesize-property-excel.md)|Returns or sets the size of the hole in a doughnut chart group. The hole size is expressed as a percentage of the chart size, between 10 and 90 percent. Read/write  **Long** .|
|[DownBars](chartgroup-downbars-property-excel.md)|Returns a  **[DownBars](downbars-object-excel.md)** object that represents the down bars on a line chart. Applies only to line charts. Read-only.|
|[DropLines](chartgroup-droplines-property-excel.md)|Returns a  **[DropLines](droplines-object-excel.md)** object that represents the drop lines for a series on a line chart or area chart. Applies only to line charts or area charts. Read-only.|
|[FirstSliceAngle](chartgroup-firstsliceangle-property-excel.md)|Returns or sets the angle of the first pie-chart or doughnut-chart slice, in degrees (clockwise from vertical). Applies only to pie, 3-D pie, and doughnut charts. Can be a value from 0 through 360. Read/write  **Long** .|
|[GapWidth](chartgroup-gapwidth-property-excel.md)|Bar and Column charts: Returns or sets the space between bar or column clusters, as a percentage of the bar or column width. Pie of Pie and Bar of Pie charts: Returns or sets the space between the primary and secondary sections of the chart. Read/write  **Long** .|
|[Has3DShading](chartgroup-has3dshading-property-excel.md)|Returns or sets the 3D Shading property of a  **ChartGroup** object. Read/write **Boolean** .|
|[HasDropLines](chartgroup-hasdroplines-property-excel.md)| **True** if the line chart or area chart has drop lines. Applies only to line and area charts. Read/write **Boolean** .|
|[HasHiLoLines](chartgroup-hashilolines-property-excel.md)| **True** if the line chart has high-low lines. Applies only to line charts. Read/write **Boolean** .|
|[HasRadarAxisLabels](chartgroup-hasradaraxislabels-property-excel.md)| **True** if a radar chart has axis labels. Applies only to radar charts. Read/write **Boolean** .|
|[HasSeriesLines](chartgroup-hasserieslines-property-excel.md)| **True** if a stacked column chart or bar chart has series lines or if a Pie of Pie chart or Bar of Pie chart has connector lines between the two sections. Applies only to 2-D stacked bar, 2-D stacked column, pie of pie, or bar of pie charts. Read/write **Boolean** .|
|[HasUpDownBars](chartgroup-hasupdownbars-property-excel.md)| **True** if a line chart has up and down bars. Applies only to line charts. Read/write **Boolean** .|
|[HiLoLines](chartgroup-hilolines-property-excel.md)|Returns a  **[HiLoLines](hilolines-object-excel.md)** object that represents the high-low lines for a series on a line chart. Applies only to line charts. Read-only.|
|[Index](chartgroup-index-property-excel.md)|Returns a  **Long** value that represents the index number of the object within the collection of similar objects.|
|[Overlap](chartgroup-overlap-property-excel.md)|Specifies how bars and columns are positioned. Can be a value between - 100 and 100. Applies only to 2-D bar and 2-D column charts. Read/write  **Long** .|
|[Parent](chartgroup-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[RadarAxisLabels](chartgroup-radaraxislabels-property-excel.md)|Returns a  **[TickLabels](ticklabels-object-excel.md)** object that represents the radar axis labels for the specified chart group. Read-only.|
|[SecondPlotSize](chartgroup-secondplotsize-property-excel.md)|Returns or sets the size of the secondary section of either a pie of pie chart or a bar of pie chart, as a percentage of the size of the primary pie. Can be a value from 5 to 200. Read/write  **Long** .|
|[SeriesLines](chartgroup-serieslines-property-excel.md)|Returns a  **[SeriesLines](serieslines-object-excel.md)** object that represents the series lines for a 2-D stacked bar, 2-D stacked column, pie of pie, or bar of pie chart. Read-only.|
|[ShowNegativeBubbles](chartgroup-shownegativebubbles-property-excel.md)| **True** if negative bubbles are shown for the chart group. Valid only for bubble charts. Read/write **Boolean** .|
|[SizeRepresents](chartgroup-sizerepresents-property-excel.md)|Returns or sets what the bubble size represents on a bubble chart. Can be either of the following  **[XlSizeRepresents](xlsizerepresents-enumeration-excel.md)** constants: **xlSizeIsArea** or **xlSizeIsWidth** . Read/write **Long** .|
|[SplitType](chartgroup-splittype-property-excel.md)|Returns or sets the way the two sections of either a pie of pie chart or a bar of pie chart are split. Read/write  **[XlChartSplitType](xlchartsplittype-enumeration-excel.md)** .|
|[SplitValue](chartgroup-splitvalue-property-excel.md)|Returns or sets the threshold value separating the two sections of either a pie of pie chart or a bar of pie chart. Read/write  **Variant** .|
|[UpBars](chartgroup-upbars-property-excel.md)|Returns an  **[UpBars](upbars-object-excel.md)** object that represents the up bars on a line chart. Applies only to line charts. Read-only.|
|[VaryByCategories](chartgroup-varybycategories-property-excel.md)| **True** if Microsoft Excel assigns a different color or pattern to each data marker. The chart must contain only one series. Read/write **Boolean** .|

