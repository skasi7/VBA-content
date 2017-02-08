---
title: Chart Members (PowerPoint)
ms.prod: POWERPOINT
ms.assetid: de1c852d-e599-3e66-1365-dde3e1eb4c28
---


# Chart Members (PowerPoint)
Represents a chart in a presentation.

Represents a chart in a presentation.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[ApplyChartTemplate](chart-applycharttemplate-method-powerpoint.md)|Applies a standard or custom chart type to a chart.|
|[ApplyDataLabels](chart-applydatalabels-method-powerpoint.md)|Applies data labels to all the series in a chart.|
|[ApplyLayout](chart-applylayout-method-powerpoint.md)|Applies the layouts shown in the Ribbon.|
|[Axes](chart-axes-method-powerpoint.md)|Returns a collection of axes on the chart.|
|[ChartGroups](chart-chartgroups-method-powerpoint.md)|Returns an object that represents either a single chart group or a collection of all the chart groups in the chart.|
|[ChartWizard](chart-chartwizard-method-powerpoint.md)|Modifies the properties of the given chart. You can use this method to quickly format a chart without setting all the individual properties. This method is noninteractive, and it changes only the specified properties.|
|[ClearToMatchColorStyle](chart-cleartomatchcolorstyle-method-powerpoint.md)|Clears all colors on the specified chart that do not follow the color style applied to the chart.|
|[ClearToMatchStyle](chart-cleartomatchstyle-method-powerpoint.md)|Clears the chart elements formatting to automatic.|
|[Copy](chart-copy-method-powerpoint.md)|Not supported for this object.|
|[CopyPicture](chart-copypicture-method-powerpoint.md)|Copies the selected object to the Clipboard as a picture.|
|[Delete](chart-delete-method-powerpoint.md)|Deletes the object.|
|[Export](chart-export-method-powerpoint.md)|Exports the chart in a graphic format.|
|[FullSeriesCollection](chart-fullseriescollection-method-powerpoint.md)|Returns the collection of all the series in the specified chart, or the specified series.|
|[GetChartElement](chart-getchartelement-method-powerpoint.md)|Returns information about the chart element at the specified x-coordinate and y-coordinate. |
|[Paste](chart-paste-method-powerpoint.md)|Pastes chart data from the Clipboard into the chart.|
|[Refresh](chart-refresh-method-powerpoint.md)|Causes the specified chart to be redrawn immediately.|
|[SaveChartTemplate](chart-savecharttemplate-method-powerpoint.md)|Saves a custom chart template to the list of available chart templates.|
|[Select](chart-select-method-powerpoint.md)|Selects the object.|
|[SeriesCollection](chart-seriescollection-method-powerpoint.md)|Returns all the series in the chart.|
|[SetBackgroundPicture](chart-setbackgroundpicture-method-powerpoint.md)|Sets the background graphic for a chart.|
|[SetDefaultChart](chart-setdefaultchart-method-powerpoint.md)|Specifies the name of the chart template that Microsoft Word uses when it creates new charts.|
|[SetElement](chart-setelement-method-powerpoint.md)|Sets chart elements on a chart. Read/write  **MsoChartElementType**.|
|[SetSourceData](chart-setsourcedata-method-powerpoint.md)|Sets the source data range for the chart.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[AlternativeText](chart-alternativetext-property-powerpoint.md)|Returns or sets the alternative text associated with a shape in a Web presentation. Read/write.|
|[Application](chart-application-property-powerpoint.md)|When used without an object qualifier, returns an  **[Application](application-object-powerpoint.md)** object that represents the Microsoft PowerPoint application. When used with an object qualifier, returns an **Application** object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.|
|[AutoScaling](chart-autoscaling-property-powerpoint.md)|**True** if Microsoft Word scales a 3-D chart so that it is closer in size to the equivalent 2-D chart. The **[RightAngleAxes](chart-rightangleaxes-property-powerpoint.md)** property must be **True**. Read/write **Boolean**.|
|[BackWall](chart-backwall-property-powerpoint.md)|Returns an object that allows the user to individually format the back wall of a 3-D chart. Read-only  **[Walls](walls-object-powerpoint.md)**.|
|[BarShape](chart-barshape-property-powerpoint.md)|Returns or sets the shape used for every series in a 3-D bar or column chart. Read/write  **[XlBarShape](xlbarshape-enumeration-powerpoint.md)**.|
|[CategoryLabelLevel](chart-categorylabellevel-property-powerpoint.md)|Returns or sets an [XlCategoryLabel](xlcategorylabellevel-enumeration-word.md) constant that specifies the source level of the chart category labels. Read-write.|
|[ChartArea](chart-chartarea-property-powerpoint.md)|Returns the complete chart area for the chart. Read-only  **[ChartArea](chartarea-object-powerpoint.md)**.|
|[ChartColor](chart-chartcolor-property-powerpoint.md)|Returns or sets an integer that represents the color scheme for the chart. Read-write.|
|[ChartData](chart-chartdata-property-powerpoint.md)|Returns information about the linked or embedded data associated with a chart. Read-only  **[ChartData](chartdata-object-powerpoint.md)**.|
|[ChartStyle](chart-chartstyle-property-powerpoint.md)|Returns or sets the chart style for the chart. Read/write  **Variant**.|
|[ChartTitle](chart-charttitle-property-powerpoint.md)|Returns the title of the specified chart. Read-only  **[ChartTitle](charttitle-object-powerpoint.md)**.|
|[ChartType](chart-charttype-property-powerpoint.md)|Returns or sets the chart type. Read/write  **[XlChartType](xlcharttype-enumeration-excel.md)**.|
|[Creator](chart-creator-property-powerpoint.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.|
|[DataTable](chart-datatable-property-powerpoint.md)|Returns the chart data table. Read-only  **[DataTable](datatable-object-powerpoint.md)**.|
|[DepthPercent](chart-depthpercent-property-powerpoint.md)|Returns or sets the depth of a 3-D chart as a percentage of the chart width (between 20 and 2000 percent). Read/write  **Long**.|
|[DisplayBlanksAs](chart-displayblanksas-property-powerpoint.md)|Returns or sets the way that blank cells are plotted on a chart. Can be one of the  **[XlDisplayBlanksAs](xldisplayblanksas-enumeration-powerpoint.md)** constants. Read/write **Long**.|
|[Elevation](chart-elevation-property-powerpoint.md)|Returns or sets the elevation, in degrees, of the 3-D chart view. Read/write  **Long**.|
|[Floor](chart-floor-property-powerpoint.md)|Returns the floor of the 3-D chart. Read-only  **[Floor](floor-object-powerpoint.md)**.|
|[Format](chart-format-property-powerpoint.md)|Returns the [ChartFormat](chartformat-object-powerpoint.md) object. Read-only.|
|[GapDepth](chart-gapdepth-property-powerpoint.md)|Returns or sets the distance, as a percentage of the marker width, between the data series in a 3-D chart. Read/write  **Long**.|
|[HasAxis](chart-hasaxis-property-powerpoint.md)|Returns or sets which axes exist on the chart. Read/write  **Variant**.|
|[HasDataTable](chart-hasdatatable-property-powerpoint.md)|**True** if the chart has a data table. Read/write **Boolean**.|
|[HasLegend](chart-haslegend-property-powerpoint.md)|**True** if the chart has a legend. Read/write **Boolean**.|
|[HasTitle](chart-hastitle-property-powerpoint.md)|**True** if the axis or chart has a visible title. Read/write **Boolean**.|
|[HeightPercent](chart-heightpercent-property-powerpoint.md)|Returns or sets the height of a 3-D chart as a percentage of the chart width (from 5 through 500 percent). Read/write  **Long**.|
|[Legend](chart-legend-property-powerpoint.md)|Returns the legend for the chart. Read-only  **[Legend](legend-object-powerpoint.md)**.|
|[Name](chart-name-property-powerpoint.md)|Read/write|
|[Parent](chart-parent-property-powerpoint.md)|Returns the parent for the specified object. Read-only  **Object**.|
|[Perspective](chart-perspective-property-powerpoint.md)|Returns or sets the perspective for the 3-D chart view. Read/write  **Long**.|
|[PlotArea](chart-plotarea-property-powerpoint.md)|Returns the plot area of a chart. Read-only  **[PlotArea](plotarea-object-powerpoint.md)**.|
|[PlotBy](chart-plotby-property-powerpoint.md)|Returns or sets the way columns or rows are used as data series on the chart. Read/write  **Long**.|
|[PlotVisibleOnly](chart-plotvisibleonly-property-powerpoint.md)|**True** if only visible cells are plotted. **False** if both visible and hidden cells are plotted. Read/write **Boolean**.|
|[RightAngleAxes](chart-rightangleaxes-property-powerpoint.md)|**True** if the chart axes are at right angles, independent of chart rotation or elevation. Read/write **Boolean**.|
|[Rotation](chart-rotation-property-powerpoint.md)|Returns or sets the rotation, in degrees, of the 3-D chart view (the rotation of the plot area around the z-axis). Read/write  **Variant**.|
|[SeriesNameLevel](chart-seriesnamelevel-property-powerpoint.md)|Returns or sets an [XlSeriesNameLevel](xlseriesnamelevel-enumeration-word.md) constant that specifies the source level of the series names. Read-write.|
|[Shapes](chart-shapes-property-powerpoint.md)|Returns a collection that represents all the shapes on the chart sheet. Read-only  **[Shapes](shapes-object-powerpoint.md)**.|
|[ShowAllFieldButtons](chart-showallfieldbuttons-property-powerpoint.md)|Returns or sets a value that indicates whether to display all field buttons on a PivotChart. Read/write.|
|[ShowAxisFieldButtons](chart-showaxisfieldbuttons-property-powerpoint.md)|Returns or sets a value that indicates whether to display axis field buttons on a PivotChart. Read/write|
|[ShowDataLabelsOverMaximum](chart-showdatalabelsovermaximum-property-powerpoint.md)|Returns or sets a value that indicates whether to show the data labels when the value is greater than the maximum value on the value axis. Read/write  **Boolean**.|
|[ShowLegendFieldButtons](chart-showlegendfieldbuttons-property-powerpoint.md)|Returns or sets a value that indicates whether to display legend field buttons on a PivotChart. Read/write.|
|[ShowReportFilterFieldButtons](chart-showreportfilterfieldbuttons-property-powerpoint.md)|Returns or sets a value that indicates whether to display the report filter field buttons on a PivotChart. Read/write.|
|[ShowValueFieldButtons](chart-showvaluefieldbuttons-property-powerpoint.md)|Returns or sets a value that indicates whether to display the value field buttons on a PivotChart. Read/write.|
|[SideWall](chart-sidewall-property-powerpoint.md)|Returns a  **[Walls](walls-object-powerpoint.md)** object that allows the user to individually format the side wall of a 3-D chart. Read-only.|
|[Title](chart-title-property-powerpoint.md)|Gets or sets a  **String** that represents the title of the chart. Read/write.|
|[Walls](chart-walls-property-powerpoint.md)|Returns the walls of the 3-D chart. Read-only  **[Walls](walls-object-powerpoint.md)**.|

