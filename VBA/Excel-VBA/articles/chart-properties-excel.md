---
title: Chart Properties (Excel)
ms.prod: EXCEL
ms.assetid: fd0c8648-eff8-4cef-a6da-22d3cdc0850a
---


# Chart Properties (Excel)

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](chart-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[AutoScaling](chart-autoscaling-property-excel.md)| **True** if Microsoft Excel scales a 3-D chart so that it's closer in size to the equivalent 2-D chart. The **[RightAngleAxes](chart-rightangleaxes-property-excel.md)** property must be **True** . Read/write **Boolean** .|
|[BackWall](chart-backwall-property-excel.md)|Returns a  **[Walls](walls-object-excel.md)** object that allows the user to individually format the back wall of a 3-D chart. Read-only.|
|[BarShape](chart-barshape-property-excel.md)|Returns or sets the shape used with the 3-D bar or column chart. Read/write  **[XlBarShape](xlbarshape-enumeration-excel.md)** .|
|[CategoryLabelLevel](chart-categorylabellevel-property-excel.md)|Returns a  **[XlCategoryLabelLevel Enumeration (Excel)](xlcategorylabellevel-enumeration-excel.md)** constant referring to the level of where the category labels are being sourced from. **Integer** Read/Write.|
|[ChartArea](chart-chartarea-property-excel.md)|Returns a  **[ChartArea](chartarea-object-excel.md)** object that represents the complete chart area for the chart. Read-only.|
|[ChartColor](chart-chartcolor-property-excel.md)|Returns or sets an  **Integer** that represents the color scheme for the chart. Read-write.|
|[ChartStyle](chart-chartstyle-property-excel.md)|Returns or sets the chart style for the chart. Read/write  **Variant** .|
|[ChartTitle](chart-charttitle-property-excel.md)|Returns a  **[ChartTitle](charttitle-object-excel.md)** object that represents the title of the specified chart. Read-only.|
|[ChartType](chart-charttype-property-excel.md)|Returns or sets the chart type. Read/write  **[XlChartType](xlcharttype-enumeration-excel.md)** .|
|[CodeName](chart-codename-property-excel.md)|Returns the code name for the object. Read-only  **String** .|
|[Creator](chart-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[DataTable](chart-datatable-property-excel.md)|Returns a  **[DataTable](datatable-object-excel.md)** object that represents the chart data table. Read-only.|
|[DepthPercent](chart-depthpercent-property-excel.md)|Returns or sets the depth of a 3-D chart as a percentage of the chart width (between 20 and 2000 percent). Read/write  **Long** .|
|[DisplayBlanksAs](chart-displayblanksas-property-excel.md)|Returns or sets the way that blank cells are plotted on a chart. Can be one of the  **[XlDisplayBlanksAs](xldisplayblanksas-enumeration-excel.md)** constants. Read/write **Long** .|
|[Elevation](chart-elevation-property-excel.md)|Returns or sets the elevation of the 3-D chart view, in degrees. Read/write  **Long** .|
|[Floor](chart-floor-property-excel.md)|Returns a  **[Floor](floor-object-excel.md)** object that represents the floor of the 3-D chart. Read-only.|
|[GapDepth](chart-gapdepth-property-excel.md)|Returns or sets the distance between the data series in a 3-D chart, as a percentage of the marker width. The value of this property must be between 0 and 500. Read/write  **Long** .|
|[HasAxis](chart-hasaxis-property-excel.md)|Returns or sets which axes exist on the chart. Read/write  **Variant** .|
|[HasDataTable](chart-hasdatatable-property-excel.md)| **True** if the chart has a data table. Read/write **Boolean** .|
|[HasLegend](chart-haslegend-property-excel.md)| **True** if the chart has a legend. Read/write **Boolean** .|
|[HasTitle](chart-hastitle-property-excel.md)| **True** if the axis or chart has a visible title. Read/write **Boolean** .|
|[HeightPercent](chart-heightpercent-property-excel.md)|Returns or sets the height of a 3-D chart as a percentage of the chart width (between 5 and 500 percent). Read/write  **Long** .|
|[Hyperlinks](chart-hyperlinks-property-excel.md)|Returns a  **[Hyperlinks](hyperlinks-object-excel.md)** collection that represents the hyperlinks for the chart.|
|[Index](chart-index-property-excel.md)|Returns a  **Long** value that represents the index number of the object within the collection of similar objects.|
|[Legend](chart-legend-property-excel.md)|Returns a  **[Legend](legend-object-excel.md)** object that represents the legend for the chart. Read-only.|
|[MailEnvelope](chart-mailenvelope-property-excel.md)|Rrepresents an e-mail header for a document.|
|[Name](chart-name-property-excel.md)|Returns or sets a  **String** value representing the name of the object.|
|[Next](chart-next-property-excel.md)|Returns a  **[Worksheet](worksheet-object-excel.md)** object that represents the next sheet.|
|[PageSetup](chart-pagesetup-property-excel.md)|Returns a  **[PageSetup](pagesetup-object-excel.md)** object that contains all the page setup settings for the specified object. Read-only.|
|[Parent](chart-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[Perspective](chart-perspective-property-excel.md)|Returns or sets a  **Long** value that represents the perspective for the 3-D chart view.|
|[PivotLayout](chart-pivotlayout-property-excel.md)|Returns a  **[PivotLayout](pivotlayout-object-excel.md)** object that represents the placement of fields in a PivotTable report and the placement of axes in a PivotChart report. Read-only.|
|[PlotArea](chart-plotarea-property-excel.md)|Returns a  **[PlotArea](plotarea-object-excel.md)** object that represents the plot area of a chart. Read-only.|
|[PlotBy](chart-plotby-property-excel.md)|Returns or sets the way columns or rows are used as data series on the chart. Can be one of the following  **[XlRowCol](xlrowcol-enumeration-excel.md)** constants: **xlColumns** or **xlRows** . Read/write **Long** .|
|[PlotVisibleOnly](chart-plotvisibleonly-property-excel.md)| **True** if only visible cells are plotted. **False** if both visible and hidden cells are plotted. Read/write **Boolean** .|
|[Previous](chart-previous-property-excel.md)|Returns a  **[Worksheet](worksheet-object-excel.md)** object that represents the next sheet.|
|[PrintedCommentPages](chart-printedcommentpages-property-excel.md)|Returns the number of comment pages that will be printed for the current chart. Read-only|
|[ProtectContents](chart-protectcontents-property-excel.md)| **True** if the contents of the sheet are protected. For a chart, this protects the entire chart. To turn on content protection, use the **[Protect](chart-protect-method-excel.md)** method with the _Contents_ argument set to **True** . Read-only **Boolean** .|
|[ProtectData](chart-protectdata-property-excel.md)| **True** if series formulas cannot be modified by the user. Read/write **Boolean** .|
|[ProtectDrawingObjects](chart-protectdrawingobjects-property-excel.md)| **True** if shapes are protected. To turn on shape protection, use the **[Protect](chart-protect-method-excel.md)** method with the _DrawingObjects_ argument set to **True** . Read-only **Boolean** .|
|[ProtectFormatting](chart-protectformatting-property-excel.md)| **True** if chart formatting cannot be modified by the user. Read/write **Boolean** .|
|[ProtectionMode](chart-protectionmode-property-excel.md)| **True** if user-interface-only protection is turned on. To turn on user interface protection, use the **[Protect](chart-protect-method-excel.md)** method with the _UserInterfaceOnly_ argument set to **True** . Read-only **Boolean** .|
|[ProtectSelection](chart-protectselection-property-excel.md)| **True** if chart elements cannot be selected. Read/write **Boolean** .|
|[RightAngleAxes](chart-rightangleaxes-property-excel.md)| **True** if the chart axes are at right angles, independent of chart rotation or elevation. Applies only to 3-D line, column, and bar charts. Read/write **Boolean** .|
|[Rotation](chart-rotation-property-excel.md)|Returns or sets the rotation of the 3-D chart view (the rotation of the plot area around the z-axis, in degrees). The value of this property must be from 0 to 360, except for 3-D bar charts, where the value must be from 0 to 44. The default value is 20. Applies only to 3-D charts. Read/write  **Variant** .|
|[SeriesNameLevel](chart-seriesnamelevel-property-excel.md)|Returns a  **[XlSeriesNameLevel Enumeration (Excel)](xlseriesnamelevel-enumeration-excel.md)** constant referring to the level of where the series names are being sourced from. **Integer** Read/Write.|
|[Shapes](chart-shapes-property-excel.md)|Returns a  **[Shapes](shapes-object-excel.md)** collection that represents all the shapes on the chart sheet. Read-only.|
|[ShowAllFieldButtons](chart-showallfieldbuttons-property-excel.md)|Returns or sets whether to display all field buttons on a PivotChart. Read/write|
|[ShowAxisFieldButtons](chart-showaxisfieldbuttons-property-excel.md)|Returns or sets whether to display axis field buttons on a PivotChart. Read/write|
|[ShowDataLabelsOverMaximum](chart-showdatalabelsovermaximum-property-excel.md)[ShowExpandCollapseEntireFieldButtons](chart-showexpandcollapseentirefieldbuttons-property-excel.md)|Returns or sets whether to show the data labels when the value is greater than the maximum value on the value axis. Read/write  **Boolean** . **True** to display the **Expand Entire Field** and **Collapse Entire Field** buttons on the specified pivot chart. Read/write **Boolean**.|
|[ShowLegendFieldButtons](chart-showlegendfieldbuttons-property-excel.md)|Returns or sets whether to display legend field buttons on a PivotChart. Read/write|
|[ShowReportFilterFieldButtons](chart-showreportfilterfieldbuttons-property-excel.md)|Returns or sets whether to display the report filter field buttons on a PivotChart. Read/write|
|[ShowValueFieldButtons](chart-showvaluefieldbuttons-property-excel.md)|Returns or sets whether to display the value field buttons on a PivotChart. Read/write|
|[SideWall](chart-sidewall-property-excel.md)|Returns a  **[Walls](walls-object-excel.md)** object that allows the user to individually format the side wall of a 3-D chart. Read-only.|
|[Tab](chart-tab-property-excel.md)|Returns a  **[Tab](tab-object-excel.md)** object for a chart.|
|[Visible](chart-visible-property-excel.md)|Returns or sets an  **[XlSheetVisibility](xlsheetvisibility-enumeration-excel.md)** value that determines whether the object is visible.|
|[Walls](chart-walls-property-excel.md)|Returns a  **[Walls](walls-object-excel.md)** object that represents the walls of the 3-D chart. Read-only.|
|||
|[ShowDataLabelsOverMaximum](chart-showdatalabelsovermaximum-property-excel.md)|Returns or sets whether to show the data labels when the value is greater than the maximum value on the value axis. Read/write  **Boolean** .|
|[ShowExpandCollapseEntireFieldButtons](chart-showexpandcollapseentirefieldbuttons-property-excel.md)| **True** to display the **Expand Entire Field** and **Collapse Entire Field** buttons on the specified pivot chart. Read/write **Boolean**.|

