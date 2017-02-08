---
title: ChartGroup Properties (Word)
ms.prod: WORD
ms.assetid: 4a71e329-4ec0-4ddf-aadd-4b070b973597
---


# ChartGroup Properties (Word)

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](chartgroup-application-property-word.md)|When used without an object qualifier, returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application. When used with an object qualifier, returns an **Application** object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.|
|[AxisGroup](chartgroup-axisgroup-property-word.md)|Returns the type of axis group. Read/write  **[XlAxisGroup](xlaxisgroup-enumeration-word.md)** .|
|[BinsCountValue](chartgroup-binscountvalue-property-word.md)|Specifies the number of bins in the histogram chart. Read/write  **Long**.|
|[BinsOverflowEnabled](chartgroup-binsoverflowenabled-property-word.md)|Specifies whether a bin for values above the [BinsOverflowValue](chartgroup-binsoverflowvalue-property-excel.md) is enabled. Read/write **Boolean**.|
|[BinsOverflowValue](chartgroup-binsoverflowvalue-property-word.md)|If an [BinsOverflowEnabled](chartgroup-binsoverflowenabled-property-excel.md) is **True**, specifies the value above which an overflow bin is displayed. Read/write  **Double**.|
|[BinsType](chartgroup-binstype-property-word.md)|Specifies how the horizontal axis of the histogram chart is formatted, by bins type. Read/write [XlBinsType](xlbinstype-enumeration-word.md).|
|[BinsUnderflowEnabled](chartgroup-binsunderflowenabled-property-word.md)|Specifies whether a bin for values below the [BinsUnderflowValue](chartgroup-binsunderflowvalue-property-word.md) is enabled. Read/write **Boolean**.|
|[BinsUnderflowValue](chartgroup-binsunderflowvalue-property-word.md)|If an [BinsUnderflowEnabled](chartgroup-binsunderflowenabled-property-word.md) is **True**, specifies the value below which an underflow bin is displayed. Read/write  **Double**.|
|[BinWidthValue](chartgroup-binwidthvalue-property-word.md)|Specifies the number of points in each range. Read/write  **Double**.|
|[BubbleScale](chartgroup-bubblescale-property-word.md)|Returns or sets the scale factor for bubbles in the specified chart group. Read/write  **Long** .|
|[Creator](chartgroup-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[DoughnutHoleSize](chartgroup-doughnutholesize-property-word.md)|Returns or sets the size of the hole in a doughnut chart group. Read/write  **Long** .|
|[DownBars](chartgroup-downbars-property-word.md)|Returns the down bars on a line chart. Read-only  **[DownBars](downbars-object-word.md)** .|
|[DropLines](chartgroup-droplines-property-word.md)|Returns the drop lines for a series on a line chart or area chart. Read-only  **[DropLines](droplines-object-word.md)** .|
|[FirstSliceAngle](chartgroup-firstsliceangle-property-word.md)|Returns or sets the angle, in degrees (clockwise from vertical), of the first pie-chart or doughnut-chart slice. Read/write  **Long** .|
|[GapWidth](chartgroup-gapwidth-property-word.md)|For bar and column charts, returns or sets the space, as a percentage of the bar or column width, between bar or column clusters. For pie-of-pie and bar-of-pie charts, returns or sets the space between the primary and secondary sections of the chart. Read/write  **Long** .|
|[Has3DShading](chartgroup-has3dshading-property-word.md)| **True** if a chart group has three-dimensional shading. Read/write **Boolean** .|
|[HasDropLines](chartgroup-hasdroplines-property-word.md)| **True** if the line chart or area chart has drop lines. Read/write **Boolean** .|
|[HasHiLoLines](chartgroup-hashilolines-property-word.md)| **True** if the line chart has high-low lines. Read/write **Boolean** .|
|[HasRadarAxisLabels](chartgroup-hasradaraxislabels-property-word.md)| **True** if a radar chart has axis labels. Read/write **Boolean** .|
|[HasSeriesLines](chartgroup-hasserieslines-property-word.md)| **True** if a stacked column chart or bar chart has series lines or if a pie-of-pie chart or bar-of-pie chart has connector lines between the two sections. Read/write **Boolean** .|
|[HasUpDownBars](chartgroup-hasupdownbars-property-word.md)| **True** if a line chart has up and down bars. Read/write **Boolean** .|
|[HiLoLines](chartgroup-hilolines-property-word.md)|Returns the high-low lines for a series on a line chart. Read-only  **[HiLoLines](hilolines-object-word.md)** .|
|[Index](chartgroup-index-property-word.md)|Returns the index number of the object within the collection of similar objects. Read-only  **Long** .|
|[Overlap](chartgroup-overlap-property-word.md)|Specifies how bars and columns are positioned. Read/write  **Long** .|
|[Parent](chartgroup-parent-property-word.md)|Returns the parent for the specified object. Read-only  **Object** .|
|[RadarAxisLabels](chartgroup-radaraxislabels-property-word.md)|Returns the radar axis labels for the specified chart group. Read-only  **[TickLabels](ticklabels-object-word.md)** .|
|[SecondPlotSize](chartgroup-secondplotsize-property-word.md)|Returns or sets the size, as a percentage of the primary pie, of the secondary section of either a pie-of-pie chart or a bar-of-pie chart. Read/write  **Long** .|
|[SeriesLines](chartgroup-serieslines-property-word.md)|Returns the series lines for a 2-D stacked bar, 2-D stacked column, pie-of-pie, or bar-of-pie chart. Read-only  **[SeriesLines](serieslines-object-word.md)** .|
|[ShowNegativeBubbles](chartgroup-shownegativebubbles-property-word.md)| **True** if negative bubbles are shown for the chart group. Read/write **Boolean** .|
|[SizeRepresents](chartgroup-sizerepresents-property-word.md)|Returns or sets what the bubble size represents on a bubble chart. Read/write  **Long** .|
|[SplitType](chartgroup-splittype-property-word.md)|Returns or sets the way the two sections of either a pie-of-pie chart or a bar-of-pie chart are split. Read/write  **[XlChartSplitType](xlchartsplittype-enumeration-word.md)** .|
|[SplitValue](chartgroup-splitvalue-property-word.md)|Returns or sets the threshold value separating the two sections of either a pie-of-pie chart or a bar-of-pie chart. Read/write  **Variant** .|
|[UpBars](chartgroup-upbars-property-word.md)|Returns the up bars on a line chart. Read-only  **[UpBars](upbars-object-word.md)** .|
|[VaryByCategories](chartgroup-varybycategories-property-word.md)| **True** if Microsoft Word assigns a different color or pattern to each data marker. Read/write **Boolean** .|
|||

