---
title: Chart Object (Excel)
keywords: vbaxl10.chm147072
f1_keywords:
- vbaxl10.chm147072
ms.prod: EXCEL
api_name:
- Excel.Chart
ms.assetid: 179c32ce-49bd-6f36-ea12-89fb5443f3ea
---


# Chart Object (Excel)

Represents a chart in a workbook.


## Remarks

The chart can be either an embedded chart (contained in a  **[ChartObject](http://msdn.microsoft.com/library/chartobject-object-excel%28Office.15%29.aspx)** object) or a separate chart sheet.

The following properties and methods for returning a  **Chart** object are described in the example section:




-  **Charts** method
    
-  **ActiveChart** property
    
-  **ActiveSheet** property
    



## Example

The  **[Charts](http://msdn.microsoft.com/library/charts-object-excel%28Office.15%29.aspx)** collection contains a **Chart** object for each chart sheet in a workbook. Use **Charts** ( _index_ ), where index is the chart-sheet index number or name, to return a single **Chart** object. The chart index number represents the position of the chart sheet on the workbook tab bar. _Charts(1)_ is the first (leftmost) chart in the workbook; _Charts(Charts.Count)_ is the last (rightmost). All chart sheets are included in the index count, even if they are hidden. The chart-sheet name is shown on the workbook tab for the chart. You can use the **[Name](http://msdn.microsoft.com/library/chartobject-name-property-excel%28Office.15%29.aspx)** property to set or return the chart name. The following example changes the color of series 1 on chart sheet 1.


```
Charts(1).SeriesCollection(1).Format.Fill.ForeColor.RGB = rgbRed
```

The following example moves the chart named Sales to the end of the active workbook.




```
Charts("Sales").Move after:=Sheets(Sheets.Count)
```

The  **Chart** object is also a member of the **[Sheets](http://msdn.microsoft.com/library/sheets-object-excel%28Office.15%29.aspx)** collection, which contains all the sheets in the workbook (both chart sheets and worksheets). Use **Sheets** ( _index_ ), where _index_ is the sheet index number or name, to return a single sheet.

When a chart is the active object, you can use the  **ActiveChart** property to refer to it. A chart sheet is active if the user has selected it or if it has been activated with the **[Activate](http://msdn.microsoft.com/library/chart-activate-method-excel%28Office.15%29.aspx)** method of the **Chart** object or the **[Activate](http://msdn.microsoft.com/library/chartobject-activate-method-excel%28Office.15%29.aspx)** method of the **ChartObject** object. The following example activates chart sheet 1 and then sets the chart type and title.




```
Charts(1).Activate 
With ActiveChart 
 .Type = xlLine 
 .HasTitle = True 
 .ChartTitle.Text = "January Sales" 
End With
```

An embedded chart is active if the user has selected it or the  **[ChartObject](http://msdn.microsoft.com/library/chartobject-object-excel%28Office.15%29.aspx)** object in which it is contained has been activated with the **[Activate](http://msdn.microsoft.com/library/chartobject-activate-method-excel%28Office.15%29.aspx)** method. The following example activates embedded chart 1 on worksheet 1 and then sets the chart type and title. Notice that after the embedded chart has been activated, the code in this example is the same as that in the previous example. Using the **ActiveChart** property allows you to write Visual Basic code that can refer to either an embedded chart or a chart sheet (whichever is active).




```
Worksheets(1).ChartObjects(1).Activate 
ActiveChart.ChartType = xlLine 
ActiveChart.HasTitle = True 
ActiveChart.ChartTitle.Text = "January Sales"
```

When a chart sheet is the active sheet, you can use the  **ActiveSheet** property to refer to it. The following example uses the **Activate** method to activate the chart sheet named Chart1 and then sets the interior color for series 1 in the chart to blue.




```
Charts("chart1").Activate 
ActiveSheet.SeriesCollection(1).Format.Fill.ForeColor.RGB = rgbBlue
```


## Events



|**Name**|
|:-----|
|[Activate](http://msdn.microsoft.com/library/chart-activate-event-excel%28Office.15%29.aspx)|
|[BeforeDoubleClick](http://msdn.microsoft.com/library/chart-beforedoubleclick-event-excel%28Office.15%29.aspx)|
|[BeforeRightClick](http://msdn.microsoft.com/library/chart-beforerightclick-event-excel%28Office.15%29.aspx)|
|[Calculate](http://msdn.microsoft.com/library/chart-calculate-event-excel%28Office.15%29.aspx)|
|[Deactivate](http://msdn.microsoft.com/library/chart-deactivate-event-excel%28Office.15%29.aspx)|
|[MouseDown](http://msdn.microsoft.com/library/chart-mousedown-event-excel%28Office.15%29.aspx)|
|[MouseMove](http://msdn.microsoft.com/library/chart-mousemove-event-excel%28Office.15%29.aspx)|
|[MouseUp](http://msdn.microsoft.com/library/chart-mouseup-event-excel%28Office.15%29.aspx)|
|[Resize](http://msdn.microsoft.com/library/chart-resize-event-excel%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/chart-select-event-excel%28Office.15%29.aspx)|
|[SeriesChange](http://msdn.microsoft.com/library/chart-serieschange-event-excel%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Activate](http://msdn.microsoft.com/library/chart-activate-method-excel%28Office.15%29.aspx)|
|[ApplyChartTemplate](http://msdn.microsoft.com/library/chart-applycharttemplate-method-excel%28Office.15%29.aspx)|
|[ApplyDataLabels](http://msdn.microsoft.com/library/chart-applydatalabels-method-excel%28Office.15%29.aspx)|
|[ApplyLayout](http://msdn.microsoft.com/library/chart-applylayout-method-excel%28Office.15%29.aspx)|
|[Axes](http://msdn.microsoft.com/library/chart-axes-method-excel%28Office.15%29.aspx)|
|[ChartGroups](http://msdn.microsoft.com/library/chart-chartgroups-method-excel%28Office.15%29.aspx)|
|[ChartObjects](http://msdn.microsoft.com/library/chart-chartobjects-method-excel%28Office.15%29.aspx)|
|[ChartWizard](http://msdn.microsoft.com/library/chart-chartwizard-method-excel%28Office.15%29.aspx)|
|[CheckSpelling](http://msdn.microsoft.com/library/chart-checkspelling-method-excel%28Office.15%29.aspx)|
|[ClearToMatchColorStyle](http://msdn.microsoft.com/library/chart-cleartomatchcolorstyle-method-excel%28Office.15%29.aspx)|
|[ClearToMatchStyle](http://msdn.microsoft.com/library/chart-cleartomatchstyle-method-excel%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/chart-copy-method-excel%28Office.15%29.aspx)|
|[CopyPicture](http://msdn.microsoft.com/library/chart-copypicture-method-excel%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/chart-delete-method-excel%28Office.15%29.aspx)|
|[Evaluate](http://msdn.microsoft.com/library/chart-evaluate-method-excel%28Office.15%29.aspx)|
|[Export](http://msdn.microsoft.com/library/chart-export-method-excel%28Office.15%29.aspx)|
|[ExportAsFixedFormat](http://msdn.microsoft.com/library/chart-exportasfixedformat-method-excel%28Office.15%29.aspx)|
|[FullSeriesCollection](http://msdn.microsoft.com/library/chart-fullseriescollection-method-excel%28Office.15%29.aspx)|
|[GetChartElement](http://msdn.microsoft.com/library/chart-getchartelement-method-excel%28Office.15%29.aspx)|
|[Location](http://msdn.microsoft.com/library/chart-location-method-excel%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/chart-move-method-excel%28Office.15%29.aspx)|
|[OLEObjects](http://msdn.microsoft.com/library/chart-oleobjects-method-excel%28Office.15%29.aspx)|
|[Paste](http://msdn.microsoft.com/library/chart-paste-method-excel%28Office.15%29.aspx)|
|[PrintOut](http://msdn.microsoft.com/library/chart-printout-method-excel%28Office.15%29.aspx)|
|[PrintPreview](http://msdn.microsoft.com/library/chart-printpreview-method-excel%28Office.15%29.aspx)|
|[Protect](http://msdn.microsoft.com/library/chart-protect-method-excel%28Office.15%29.aspx)|
|[Refresh](http://msdn.microsoft.com/library/chart-refresh-method-excel%28Office.15%29.aspx)|
|[SaveAs](http://msdn.microsoft.com/library/chart-saveas-method-excel%28Office.15%29.aspx)|
|[SaveChartTemplate](http://msdn.microsoft.com/library/chart-savecharttemplate-method-excel%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/chart-select-method-excel%28Office.15%29.aspx)|
|[SeriesCollection](http://msdn.microsoft.com/library/chart-seriescollection-method-excel%28Office.15%29.aspx)|
|[SetBackgroundPicture](http://msdn.microsoft.com/library/chart-setbackgroundpicture-method-excel%28Office.15%29.aspx)|
|[SetDefaultChart](http://msdn.microsoft.com/library/chart-setdefaultchart-method-excel%28Office.15%29.aspx)|
|[SetElement](http://msdn.microsoft.com/library/chart-setelement-method-excel%28Office.15%29.aspx)|
|[SetSourceData](http://msdn.microsoft.com/library/chart-setsourcedata-method-excel%28Office.15%29.aspx)|
|[Unprotect](http://msdn.microsoft.com/library/chart-unprotect-method-excel%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/chart-application-property-excel%28Office.15%29.aspx)|
|[AutoScaling](http://msdn.microsoft.com/library/chart-autoscaling-property-excel%28Office.15%29.aspx)|
|[BackWall](http://msdn.microsoft.com/library/chart-backwall-property-excel%28Office.15%29.aspx)|
|[BarShape](http://msdn.microsoft.com/library/chart-barshape-property-excel%28Office.15%29.aspx)|
|[CategoryLabelLevel](http://msdn.microsoft.com/library/chart-categorylabellevel-property-excel%28Office.15%29.aspx)|
|[ChartArea](http://msdn.microsoft.com/library/chart-chartarea-property-excel%28Office.15%29.aspx)|
|[ChartColor](http://msdn.microsoft.com/library/chart-chartcolor-property-excel%28Office.15%29.aspx)|
|[ChartStyle](http://msdn.microsoft.com/library/chart-chartstyle-property-excel%28Office.15%29.aspx)|
|[ChartTitle](http://msdn.microsoft.com/library/chart-charttitle-property-excel%28Office.15%29.aspx)|
|[ChartType](http://msdn.microsoft.com/library/chart-charttype-property-excel%28Office.15%29.aspx)|
|[CodeName](http://msdn.microsoft.com/library/chart-codename-property-excel%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/chart-creator-property-excel%28Office.15%29.aspx)|
|[DataTable](http://msdn.microsoft.com/library/chart-datatable-property-excel%28Office.15%29.aspx)|
|[DepthPercent](http://msdn.microsoft.com/library/chart-depthpercent-property-excel%28Office.15%29.aspx)|
|[DisplayBlanksAs](http://msdn.microsoft.com/library/chart-displayblanksas-property-excel%28Office.15%29.aspx)|
|[Elevation](http://msdn.microsoft.com/library/chart-elevation-property-excel%28Office.15%29.aspx)|
|[Floor](http://msdn.microsoft.com/library/chart-floor-property-excel%28Office.15%29.aspx)|
|[GapDepth](http://msdn.microsoft.com/library/chart-gapdepth-property-excel%28Office.15%29.aspx)|
|[HasAxis](http://msdn.microsoft.com/library/chart-hasaxis-property-excel%28Office.15%29.aspx)|
|[HasDataTable](http://msdn.microsoft.com/library/chart-hasdatatable-property-excel%28Office.15%29.aspx)|
|[HasLegend](http://msdn.microsoft.com/library/chart-haslegend-property-excel%28Office.15%29.aspx)|
|[HasTitle](http://msdn.microsoft.com/library/chart-hastitle-property-excel%28Office.15%29.aspx)|
|[HeightPercent](http://msdn.microsoft.com/library/chart-heightpercent-property-excel%28Office.15%29.aspx)|
|[Hyperlinks](http://msdn.microsoft.com/library/chart-hyperlinks-property-excel%28Office.15%29.aspx)|
|[Index](http://msdn.microsoft.com/library/chart-index-property-excel%28Office.15%29.aspx)|
|[Legend](http://msdn.microsoft.com/library/chart-legend-property-excel%28Office.15%29.aspx)|
|[MailEnvelope](http://msdn.microsoft.com/library/chart-mailenvelope-property-excel%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/chart-name-property-excel%28Office.15%29.aspx)|
|[Next](http://msdn.microsoft.com/library/chart-next-property-excel%28Office.15%29.aspx)|
|[PageSetup](http://msdn.microsoft.com/library/chart-pagesetup-property-excel%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/chart-parent-property-excel%28Office.15%29.aspx)|
|[Perspective](http://msdn.microsoft.com/library/chart-perspective-property-excel%28Office.15%29.aspx)|
|[PivotLayout](http://msdn.microsoft.com/library/chart-pivotlayout-property-excel%28Office.15%29.aspx)|
|[PlotArea](http://msdn.microsoft.com/library/chart-plotarea-property-excel%28Office.15%29.aspx)|
|[PlotBy](http://msdn.microsoft.com/library/chart-plotby-property-excel%28Office.15%29.aspx)|
|[PlotVisibleOnly](http://msdn.microsoft.com/library/chart-plotvisibleonly-property-excel%28Office.15%29.aspx)|
|[Previous](http://msdn.microsoft.com/library/chart-previous-property-excel%28Office.15%29.aspx)|
|[PrintedCommentPages](http://msdn.microsoft.com/library/chart-printedcommentpages-property-excel%28Office.15%29.aspx)|
|[ProtectContents](http://msdn.microsoft.com/library/chart-protectcontents-property-excel%28Office.15%29.aspx)|
|[ProtectData](http://msdn.microsoft.com/library/chart-protectdata-property-excel%28Office.15%29.aspx)|
|[ProtectDrawingObjects](http://msdn.microsoft.com/library/chart-protectdrawingobjects-property-excel%28Office.15%29.aspx)|
|[ProtectFormatting](http://msdn.microsoft.com/library/chart-protectformatting-property-excel%28Office.15%29.aspx)|
|[ProtectionMode](http://msdn.microsoft.com/library/chart-protectionmode-property-excel%28Office.15%29.aspx)|
|[ProtectSelection](http://msdn.microsoft.com/library/chart-protectselection-property-excel%28Office.15%29.aspx)|
|[RightAngleAxes](http://msdn.microsoft.com/library/chart-rightangleaxes-property-excel%28Office.15%29.aspx)|
|[Rotation](http://msdn.microsoft.com/library/chart-rotation-property-excel%28Office.15%29.aspx)|
|[SeriesNameLevel](http://msdn.microsoft.com/library/chart-seriesnamelevel-property-excel%28Office.15%29.aspx)|
|[Shapes](http://msdn.microsoft.com/library/chart-shapes-property-excel%28Office.15%29.aspx)|
|[ShowAllFieldButtons](http://msdn.microsoft.com/library/chart-showallfieldbuttons-property-excel%28Office.15%29.aspx)|
|[ShowAxisFieldButtons](http://msdn.microsoft.com/library/chart-showaxisfieldbuttons-property-excel%28Office.15%29.aspx)|
|[ShowDataLabelsOverMaximum](http://msdn.microsoft.com/library/chart-showdatalabelsovermaximum-property-excel%28Office.15%29.aspx)[ShowExpandCollapseEntireFieldButtons](http://msdn.microsoft.com/library/chart-showexpandcollapseentirefieldbuttons-property-excel%28Office.15%29.aspx)|
|[ShowLegendFieldButtons](http://msdn.microsoft.com/library/chart-showlegendfieldbuttons-property-excel%28Office.15%29.aspx)|
|[ShowReportFilterFieldButtons](http://msdn.microsoft.com/library/chart-showreportfilterfieldbuttons-property-excel%28Office.15%29.aspx)|
|[ShowValueFieldButtons](http://msdn.microsoft.com/library/chart-showvaluefieldbuttons-property-excel%28Office.15%29.aspx)|
|[SideWall](http://msdn.microsoft.com/library/chart-sidewall-property-excel%28Office.15%29.aspx)|
|[Tab](http://msdn.microsoft.com/library/chart-tab-property-excel%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/chart-visible-property-excel%28Office.15%29.aspx)|
|[Walls](http://msdn.microsoft.com/library/chart-walls-property-excel%28Office.15%29.aspx)|
||
|[ShowDataLabelsOverMaximum](http://msdn.microsoft.com/library/chart-showdatalabelsovermaximum-property-excel%28Office.15%29.aspx)|
|[ShowExpandCollapseEntireFieldButtons](http://msdn.microsoft.com/library/chart-showexpandcollapseentirefieldbuttons-property-excel%28Office.15%29.aspx)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/object-model-excel-vba-reference%28Office.15%29.aspx)
<<<<<<< HEAD
=======

>>>>>>> d7667e83d23dbf8ebf5bf068ba6fed14c840c0f5

