---
title: ChartData Object (PowerPoint)
keywords: vbapp10.chm689000
f1_keywords:
- vbapp10.chm689000
ms.prod: POWERPOINT
ms.assetid: b7bedf0e-5f11-001d-a97c-e8d07939bc8b
---


# ChartData Object (PowerPoint)

Represents access to the linked or embedded data associated with a chart.


## Remarks

Use the  **[ChartData](http://msdn.microsoft.com/library/chart-chartdata-property-powerpoint%28Office.15%29.aspx)** property to return the **ChartData** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example uses the  **[Activate](http://msdn.microsoft.com/library/chartdata-activate-method-powerpoint%28Office.15%29.aspx)** method to display the data associated with the first chart in the active document.




```
With ActiveDocument.InlineShapes(1).Chart.ChartData

    .Activate

End With
```


## Methods



|**Name**|
|:-----|
|[Activate](http://msdn.microsoft.com/library/chartdata-activate-method-powerpoint%28Office.15%29.aspx)|
|[ActivateChartDataWindow](http://msdn.microsoft.com/library/chartdata-activatechartdatawindow-method-powerpoint%28Office.15%29.aspx)|
|[BreakLink](http://msdn.microsoft.com/library/chartdata-breaklink-method-powerpoint%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[IsLinked](http://msdn.microsoft.com/library/chartdata-islinked-property-powerpoint%28Office.15%29.aspx)|
|[Workbook](http://msdn.microsoft.com/library/chartdata-workbook-property-powerpoint%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/object-model-powerpoint-vba-reference%28Office.15%29.aspx)
