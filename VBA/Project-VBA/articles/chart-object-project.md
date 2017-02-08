---
title: Chart Object (Project)
ms.prod: PROJECTSERVER
ms.assetid: 810d4ec1-69d2-c432-b9da-57042b783b85
---


# Chart Object (Project)
The  **Chart** object represents a chart on a report in Project.




## Remarks

The  **Chart** object in Project includes the standard members that other Office applications implement for Office Art. For example, see the **Chart** object in the VBA object model for Word, Excel, and PowerPoint.

In Project, a chart is represented by a  **Chart** object, which is contained by a **[Shape](http://msdn.microsoft.com/library/shape-object-project%28Office.15%29.aspx)** object or a **[ShapeRange](http://msdn.microsoft.com/library/shaperange-object-project%28Office.15%29.aspx)** collection in a **[Report](http://msdn.microsoft.com/library/report-object-project%28Office.15%29.aspx)** object. For a diagram that shows the **Chart** object in the Project object model hierarchy, see[Application and Projects object map](http://msdn.microsoft.com/library/application-and-projects-object-map-project%28Office.15%29.aspx).


 **Note**  Macro recording for the  **Chart** object is not implemented. That is, when you record a macro in Project and manually add a chart, add chart elements, or manually format a chart in a report, the steps for adding and manipulating the chart are not recorded.

You can use the  **[Shapes.AddChart](http://msdn.microsoft.com/library/shapes-addchart-method-project%28Office.15%29.aspx)** method to add a chart to a report. To determine whether a **Shape** or a **ShapeRange** contains a chart, use the **HasChart** method.

The  **Chart** object in Project does not implement events. So, a chart in Project cannot be animated to interact with mouse events or respond to events such as **Select** or **Calculate**, as it can in Excel.


## Example

The following example creates a simple scalar chart for tasks in the active project. The chart shows the  **Actual Work**,  **Remaining Work**, and  **Work** default fields.

To create some sample data, add four tasks to a new project, assign local resources to those tasks, and set various values of duration and actual work. For example, try the values in Table 1.


**Table 1. Sample data for a simple chart**


|**Task name**|**Duration**|**Actual work**|
|:-----|:-----|:-----|
|T1|2d|16|
|T2|5d|19|
|T3|4d|7|
|T4|2d|0|
?




```
Sub AddSimpleScalarChart()
    Dim chartReport As Report
    Dim reportName As String
    
    ' Add a report.
    reportName = "Simple scalar chart"
    Set chartReport = ActiveProject.Reports.Add(reportName)

    ' Add a chart.
    Dim chartShape As Shape
    Set chartShape = ActiveProject.Reports(reportName).Shapes.AddChart()
    
    chartShape.Chart.SetElement (msoElementChartTitleCenteredOverlay)
    chartShape.Chart.ChartTitle.Text = "Sample Chart for the Test1 project"
End Sub
```

When you run the  **AddSimpleScalarChart** macro, Project creates the report and adds a chart. The chart has default features, except the title is specified by the **SetElement** property to be overlaid on the chart, instead of the default position above the chart.


**Figure 1. The chart shows the data in Table 1**

![Simple scalar chart in a report](images/pj15_VBA_ChartObject.gif)To delete the chart, you can delete the shape that contains the chart. The following macro deletes the chart on the report that is created by the  **AddSimpleScalarChart** macro, and leaves the empty report as the active view.




```
Sub DeleteTheShape()
    Dim i As Integer
    Dim reportName As String
    Dim theShape As MSProject.Shape
    
    reportName = "Simple scalar chart"
        
    For i = 1 To ActiveProject.Reports.Count
        If ActiveProject.Reports(i).Name = reportName Then
            Set theShape = ActiveProject.Reports(i).Shapes(1)
            theShape.Delete
        End If
    Next i
End Sub
```

To delete the report, go to a different view, and then open the  **Organizer** dialog box. You cannot delete a report while the report is active. The **Organizer** is available on the **DEVELOPER** tab of the ribbon, and also on the **DESIGN** tab, in the **Report** group, on the **Manage** menu. On the **Reports** tab of the **Organizer** dialog box, select **Simple scalar chart** in the project pane, and then choose **Delete**. Alternately, run the following macro to delete the report.




```
Sub DeleteTheReport()
    Dim i As Integer
    Dim reportName As String
    
    reportName = "Simple scalar chart"

    ' To delete the active report, change to another view.
    ViewApplyEx Name:="&amp;Gantt Chart"
    
    ActiveProject.Reports(reportName).Delete
End Sub
```


## Methods



|**Name**|
|:-----|
|[ApplyChartTemplate](http://msdn.microsoft.com/library/chart-applycharttemplate-method-project%28Office.15%29.aspx)|
|[ApplyCustomType](http://msdn.microsoft.com/library/chart-applycustomtype-method-project%28Office.15%29.aspx)|
|[ApplyDataLabels](http://msdn.microsoft.com/library/chart-applydatalabels-method-project%28Office.15%29.aspx)|
|[ApplyLayout](http://msdn.microsoft.com/library/chart-applylayout-method-project%28Office.15%29.aspx)|
|[AutoFormat](http://msdn.microsoft.com/library/chart-autoformat-method-project%28Office.15%29.aspx)|
|[Axes](http://msdn.microsoft.com/library/chart-axes-method-project%28Office.15%29.aspx)|
|[ChartWizard](http://msdn.microsoft.com/library/chart-chartwizard-method-project%28Office.15%29.aspx)|
|[ClearToMatchColorStyle](http://msdn.microsoft.com/library/chart-cleartomatchcolorstyle-method-project%28Office.15%29.aspx)|
|[ClearToMatchStyle](http://msdn.microsoft.com/library/chart-cleartomatchstyle-method-project%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/chart-copy-method-project%28Office.15%29.aspx)|
|[CopyPicture](http://msdn.microsoft.com/library/chart-copypicture-method-project%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/chart-delete-method-project%28Office.15%29.aspx)|
|[Export](http://msdn.microsoft.com/library/chart-export-method-project%28Office.15%29.aspx)|
|[GetChartElement](http://msdn.microsoft.com/library/chart-getchartelement-method-project%28Office.15%29.aspx)|
|[Refresh](http://msdn.microsoft.com/library/chart-refresh-method-project%28Office.15%29.aspx)|
|[RefreshPivotTable](http://msdn.microsoft.com/library/chart-refreshpivottable-method-project%28Office.15%29.aspx)|
|[SaveChartTemplate](http://msdn.microsoft.com/library/chart-savecharttemplate-method-project%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/chart-select-method-project%28Office.15%29.aspx)|
|[SeriesCollection](http://msdn.microsoft.com/library/chart-seriescollection-method-project%28Office.15%29.aspx)|
|[SetDefaultChart](http://msdn.microsoft.com/library/chart-setdefaultchart-method-project%28Office.15%29.aspx)|
|[SetElement](http://msdn.microsoft.com/library/chart-setelement-method-project%28Office.15%29.aspx)|
|[SetSourceData](http://msdn.microsoft.com/library/chart-setsourcedata-method-project%28Office.15%29.aspx)|
|[UpdateChartData](http://msdn.microsoft.com/library/chart-updatechartdata-method-project%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/chart-application-property-project%28Office.15%29.aspx)|
|[AutoScaling](http://msdn.microsoft.com/library/chart-autoscaling-property-project%28Office.15%29.aspx)|
|[BackWall](http://msdn.microsoft.com/library/chart-backwall-property-project%28Office.15%29.aspx)|
|[BarShape](http://msdn.microsoft.com/library/chart-barshape-property-project%28Office.15%29.aspx)|
|[ChartArea](http://msdn.microsoft.com/library/chart-chartarea-property-project%28Office.15%29.aspx)|
|[ChartColor](http://msdn.microsoft.com/library/chart-chartcolor-property-project%28Office.15%29.aspx)|
|[ChartData](http://msdn.microsoft.com/library/chart-chartdata-property-project%28Office.15%29.aspx)|
|[ChartGroups](http://msdn.microsoft.com/library/chart-chartgroups-property-project%28Office.15%29.aspx)|
|[ChartStyle](http://msdn.microsoft.com/library/chart-chartstyle-property-project%28Office.15%29.aspx)|
|[ChartTitle](http://msdn.microsoft.com/library/chart-charttitle-property-project%28Office.15%29.aspx)|
|[ChartType](http://msdn.microsoft.com/library/chart-charttype-property-project%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/chart-creator-property-project%28Office.15%29.aspx)|
|[DataTable](http://msdn.microsoft.com/library/chart-datatable-property-project%28Office.15%29.aspx)|
|[DepthPercent](http://msdn.microsoft.com/library/chart-depthpercent-property-project%28Office.15%29.aspx)|
|[DisplayBlanksAs](http://msdn.microsoft.com/library/chart-displayblanksas-property-project%28Office.15%29.aspx)|
|[Elevation](http://msdn.microsoft.com/library/chart-elevation-property-project%28Office.15%29.aspx)|
|[Floor](http://msdn.microsoft.com/library/chart-floor-property-project%28Office.15%29.aspx)|
|[Format](http://msdn.microsoft.com/library/chart-format-property-project%28Office.15%29.aspx)|
|[GapDepth](http://msdn.microsoft.com/library/chart-gapdepth-property-project%28Office.15%29.aspx)|
|[HasAxis](http://msdn.microsoft.com/library/chart-hasaxis-property-project%28Office.15%29.aspx)|
|[HasDataTable](http://msdn.microsoft.com/library/chart-hasdatatable-property-project%28Office.15%29.aspx)|
|[HasLegend](http://msdn.microsoft.com/library/chart-haslegend-property-project%28Office.15%29.aspx)|
|[HasTitle](http://msdn.microsoft.com/library/chart-hastitle-property-project%28Office.15%29.aspx)|
|[HeightPercent](http://msdn.microsoft.com/library/chart-heightpercent-property-project%28Office.15%29.aspx)|
|[Legend](http://msdn.microsoft.com/library/chart-legend-property-project%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/chart-parent-property-project%28Office.15%29.aspx)|
|[Perspective](http://msdn.microsoft.com/library/chart-perspective-property-project%28Office.15%29.aspx)|
|[PivotLayout](http://msdn.microsoft.com/library/chart-pivotlayout-property-project%28Office.15%29.aspx)|
|[PlotArea](http://msdn.microsoft.com/library/chart-plotarea-property-project%28Office.15%29.aspx)|
|[PlotBy](http://msdn.microsoft.com/library/chart-plotby-property-project%28Office.15%29.aspx)|
|[PlotVisibleOnly](http://msdn.microsoft.com/library/chart-plotvisibleonly-property-project%28Office.15%29.aspx)|
|[RightAngleAxes](http://msdn.microsoft.com/library/chart-rightangleaxes-property-project%28Office.15%29.aspx)|
|[Rotation](http://msdn.microsoft.com/library/chart-rotation-property-project%28Office.15%29.aspx)|
|[Shapes](http://msdn.microsoft.com/library/chart-shapes-property-project%28Office.15%29.aspx)|
|[ShowAllFieldButtons](http://msdn.microsoft.com/library/chart-showallfieldbuttons-property-project%28Office.15%29.aspx)|
|[ShowAxisFieldButtons](http://msdn.microsoft.com/library/chart-showaxisfieldbuttons-property-project%28Office.15%29.aspx)|
|[ShowDataLabelsOverMaximum](http://msdn.microsoft.com/library/chart-showdatalabelsovermaximum-property-project%28Office.15%29.aspx)|
|[ShowLegendFieldButtons](http://msdn.microsoft.com/library/chart-showlegendfieldbuttons-property-project%28Office.15%29.aspx)|
|[ShowReportFilterFieldButtons](http://msdn.microsoft.com/library/chart-showreportfilterfieldbuttons-property-project%28Office.15%29.aspx)|
|[ShowValueFieldButtons](http://msdn.microsoft.com/library/chart-showvaluefieldbuttons-property-project%28Office.15%29.aspx)|
|[SideWall](http://msdn.microsoft.com/library/chart-sidewall-property-project%28Office.15%29.aspx)|
|[Walls](http://msdn.microsoft.com/library/chart-walls-property-project%28Office.15%29.aspx)|

## See also


#### Other resources


[Shape Object](http://msdn.microsoft.com/library/shape-object-project%28Office.15%29.aspx)
