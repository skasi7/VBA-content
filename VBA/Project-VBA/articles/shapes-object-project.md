---
title: Shapes Object (Project)
ms.prod: PROJECTSERVER
ms.assetid: 6e42040c-dd5a-de4c-afa8-f9e33d1e5054
---


# Shapes Object (Project)
Represents a collection of  **[Shape](http://msdn.microsoft.com/library/shape-object-project%28Office.15%29.aspx)** objects in a custom report.

## Example

Use the  **[Report.Shapes](http://msdn.microsoft.com/library/report-shapes-property-project%28Office.15%29.aspx)** property to get the **Shapes** collection object. In the following example, the report must be the active view to get the **Shapes** collection; otherwise, you get a run-time error 424 (Object required) in the `For Each oShape In oReport.Shapes` statement.


```
Sub ListShapesInReport()
    Dim oReports As Reports
    Dim oReport As Report
    Dim oShape As shape
    Dim reportName As String
    Dim msg As String
    Dim msgBoxTitle As String
    Dim numShapes As Integer
    
    numShapes = 0
    msg = ""
    reportName = "Table Tests"
    Set oReports = ActiveProject.Reports
    
    If oReports.IsPresent(reportName) Then
        ' Make the report the active view.
        oReports(reportName).Apply
        
        Set oReport = oReports(reportName)
        msgBoxTitle = "Shapes in report: '" &amp; oReport.Name &amp; "'"
    
        For Each oShape In oReport.Shapes
            numShapes = numShapes + 1
            msg = msg &amp; numShapes &amp; ". Shape type: " &amp; CStr(oShape.Type) _
                &amp; ", '" &amp; oShape.Name &amp; "'" &amp; vbCrLf
        Next oShape
        
        If numShapes > 0 Then
            MsgBox Prompt:=msg, Title:=msgBoxTitle
        Else
            MsgBox Prompt:="This report contains no shapes.", _
                Title:=msgBoxTitle
        End If
    Else
         MsgBox Prompt:="The requested report, '" &amp; reportName _
            &amp; "', does not exist.", Title:="Report error"
    End If
End Sub
```


## Methods



|**Name**|
|:-----|
|[AddCallout](http://msdn.microsoft.com/library/shapes-addcallout-method-project%28Office.15%29.aspx)|
|[AddChart](http://msdn.microsoft.com/library/shapes-addchart-method-project%28Office.15%29.aspx)|
|[AddConnector](http://msdn.microsoft.com/library/shapes-addconnector-method-project%28Office.15%29.aspx)|
|[AddCurve](http://msdn.microsoft.com/library/shapes-addcurve-method-project%28Office.15%29.aspx)|
|[AddLabel](http://msdn.microsoft.com/library/shapes-addlabel-method-project%28Office.15%29.aspx)|
|[AddLine](http://msdn.microsoft.com/library/shapes-addline-method-project%28Office.15%29.aspx)|
|[AddPolyline](http://msdn.microsoft.com/library/shapes-addpolyline-method-project%28Office.15%29.aspx)|
|[AddShape](http://msdn.microsoft.com/library/shapes-addshape-method-project%28Office.15%29.aspx)|
|[AddTable](http://msdn.microsoft.com/library/shapes-addtable-method-project%28Office.15%29.aspx)|
|[AddTextbox](http://msdn.microsoft.com/library/shapes-addtextbox-method-project%28Office.15%29.aspx)|
|[AddTextEffect](http://msdn.microsoft.com/library/shapes-addtexteffect-method-project%28Office.15%29.aspx)|
|[BuildFreeform](http://msdn.microsoft.com/library/shapes-buildfreeform-method-project%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/shapes-item-method-project%28Office.15%29.aspx)|
|[Range](http://msdn.microsoft.com/library/shapes-range-method-project%28Office.15%29.aspx)|
|[SelectAll](http://msdn.microsoft.com/library/shapes-selectall-method-project%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Background](http://msdn.microsoft.com/library/shapes-background-property-project%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/shapes-count-property-project%28Office.15%29.aspx)|
|[Default](http://msdn.microsoft.com/library/shapes-default-property-project%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/shapes-parent-property-project%28Office.15%29.aspx)|
|[Value](http://msdn.microsoft.com/library/shapes-value-property-project%28Office.15%29.aspx)|

## See also


#### Other resources


[Shape Object](http://msdn.microsoft.com/library/shape-object-project%28Office.15%29.aspx)
[Report Object](http://msdn.microsoft.com/library/report-object-project%28Office.15%29.aspx)
[ShapeRange Object](http://msdn.microsoft.com/library/shaperange-object-project%28Office.15%29.aspx)
