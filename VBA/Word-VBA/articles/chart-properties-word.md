---
title: Chart Properties (Word)
ms.prod: WORD
ms.assetid: 795ba33a-4663-4c2a-bd63-3fc85d216616
---


# Chart Properties (Word)

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](chart-application-property-word.md)|When used without an object qualifier, returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application. When used with an object qualifier, returns an **Application** object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.|
|[AutoScaling](chart-autoscaling-property-word.md)| **True** if Microsoft Word scales a 3-D chart so that it is closer in size to the equivalent 2-D chart. The **[RightAngleAxes](chart-rightangleaxes-property-word.md)** property must be **True** . Read/write **Boolean** .|
|[BackWall](chart-backwall-property-word.md)|Returns an object that allows the user to individually format the back wall of a 3-D chart. Read-only  **[Walls](walls-object-word.md)** .|
|[BarShape](chart-barshape-property-word.md)|Returns or sets the shape used for every series in a 3-D bar or column chart. Read/write  **[XlBarShape](xlbarshape-enumeration-word.md)** .|
|[CategoryLabelLevel](chart-categorylabellevel-property-word.md)|Returns or sets an [XlCategoryLabel](xlcategorylabellevel-enumeration-word.md) constant that specifies the source level of the chart category labels. Read-write.|
|[ChartArea](chart-chartarea-property-word.md)|Returns the complete chart area for the chart. Read-only  **[ChartArea](chartarea-object-word.md)** .|
|[ChartColor](chart-chartcolor-property-word.md)|Returns or sets an integer that represents the color scheme for the chart. Read-write.|
|[ChartData](chart-chartdata-property-word.md)|Returns information about the linked or embedded data associated with a chart. Read-only  **[ChartData](chartdata-object-word.md)** .|
|[ChartGroups](chart-chartgroups-property-word.md)|Returns an object that represents either a single chart group or a collection of all the chart groups in the chart.|
|[ChartStyle](chart-chartstyle-property-word.md)|Returns or sets the chart style for the chart. Read/write  **Variant** .|
|[ChartTitle](chart-charttitle-property-word.md)|Returns the title of the specified chart. Read-only  **[ChartTitle](charttitle-object-word.md)** .|
|[ChartType](chart-charttype-property-word.md)|Returns or sets the chart type. Read/write  **[XlChartType](xlcharttype-enumeration-excel.md)** .|
|[Creator](chart-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[DataTable](chart-datatable-property-word.md)|Returns the chart data table. Read-only  **[DataTable](datatable-object-word.md)** .|
|[DepthPercent](chart-depthpercent-property-word.md)|Returns or sets the depth of a 3-D chart as a percentage of the chart width (between 20 and 2000 percent). Read/write  **Long** .|
|[DisplayBlanksAs](chart-displayblanksas-property-word.md)|Returns or sets the way that blank cells are plotted on a chart. Can be one of the  **[XlDisplayBlanksAs](xldisplayblanksas-enumeration-word.md)** constants. Read/write **Long** .|
|[Elevation](chart-elevation-property-word.md)|Returns or sets the elevation, in degrees, of the 3-D chart view. Read/write  **Long** .|
|[Floor](chart-floor-property-word.md)|Returns the floor of the 3-D chart. Read-only  **[Floor](floor-object-word.md)** .|
|[GapDepth](chart-gapdepth-property-word.md)|Returns or sets the distance, as a percentage of the marker width, between the data series in a 3-D chart. Read/write  **Long** .|
|[HasAxis](chart-hasaxis-property-word.md)|Returns or sets which axes exist on the chart. Read/write  **Variant** .|
|[HasDataTable](chart-hasdatatable-property-word.md)| **True** if the chart has a data table. Read/write **Boolean** .|
|[HasLegend](chart-haslegend-property-word.md)| **True** if the chart has a legend. Read/write **Boolean** .|
|[HasTitle](chart-hastitle-property-word.md)| **True** if the axis or chart has a visible title. Read/write **Boolean** .|
|[HeightPercent](chart-heightpercent-property-word.md)|Returns or sets the height of a 3-D chart as a percentage of the chart width (from 5 through 500 percent). Read/write  **Long** .|
|[Legend](chart-legend-property-word.md)|Returns the legend for the chart. Read-only  **[Legend](legend-object-word.md)** .|
|[Parent](chart-parent-property-word.md)|Returns the parent for the specified object. Read-only  **Object** .|
|[Perspective](chart-perspective-property-word.md)|Returns or sets the perspective for the 3-D chart view. Read/write  **Long** .|
|[PivotLayout](chart-pivotlayout-property-word.md)|Not supported for this object.|
|[PlotArea](chart-plotarea-property-word.md)|Returns the plot area of a chart. Read-only  **[PlotArea](plotarea-object-word.md)** .|
|[PlotBy](chart-plotby-property-word.md)|Returns or sets the way columns or rows are used as data series on the chart. Read/write  **Long** .|
|[PlotVisibleOnly](chart-plotvisibleonly-property-word.md)| **True** if only visible cells are plotted. **False** if both visible and hidden cells are plotted. Read/write **Boolean** .|
|[RightAngleAxes](chart-rightangleaxes-property-word.md)| **True** if the chart axes are at right angles, independent of chart rotation or elevation. Read/write **Boolean** .|
|[Rotation](chart-rotation-property-word.md)|Returns or sets the rotation, in degrees, of the 3-D chart view (the rotation of the plot area around the z-axis). Read/write  **Variant** .|
|[SeriesNameLevel](chart-seriesnamelevel-property-word.md)|Returns or sets an [XlSeriesNameLevel](xlseriesnamelevel-enumeration-word.md) constant that specifies the source level of the series names. Read-write.|
|[Shapes](chart-shapes-property-word.md)|Returns a collection that represents all the shapes on the chart sheet. Read-only  **[Shapes](shapes-object-word.md)** .|
|[ShowAllFieldButtons](chart-showallfieldbuttons-property-word.md)|Returns or sets whether to display all field buttons on a PivotChart. Read/write. Deprecated.|
|[ShowAxisFieldButtons](chart-showaxisfieldbuttons-property-word.md)|Returns or sets whether to display axis field buttons on a PivotChart. Read/write. Depcrecated.|
|[ShowDataLabelsOverMaximum](chart-showdatalabelsovermaximum-property-word.md)|Returns or sets a value that indicates whether to show the data labels when the value is greater than the maximum value on the value axis. Read/write  **Boolean** .|
|[ShowLegendFieldButtons](chart-showlegendfieldbuttons-property-word.md)|Returns or sets whether to display legend field buttons on a PivotChart. Read/write. Deprecated.|
|[ShowReportFilterFieldButtons](chart-showreportfilterfieldbuttons-property-word.md)|Returns or sets whether to display the report filter field buttons on a PivotChart. Read/write. Deprecated.|
|[ShowValueFieldButtons](chart-showvaluefieldbuttons-property-word.md)|Returns or sets whether to display the value field buttons on a PivotChart. Read/write. Deprecated.|
|[SideWall](chart-sidewall-property-word.md)|Returns a  **[Walls](walls-object-word.md)** object that allows the user to individually format the side wall of a 3-D chart. Read-only.|
|[Walls](chart-walls-property-word.md)|Returns the walls of the 3-D chart. Read-only  **[Walls](walls-object-word.md)** .|

