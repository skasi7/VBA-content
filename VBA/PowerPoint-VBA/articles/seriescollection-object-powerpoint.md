---
title: SeriesCollection Object (PowerPoint)
keywords: vbapp10.chm717000
f1_keywords:
- vbapp10.chm717000
ms.prod: POWERPOINT
ms.assetid: 6277f9e0-0198-0773-9c54-f2d009c0ba7a
---


# SeriesCollection Object (PowerPoint)

Represents a collection of all the  **[Series](series-object-powerpoint.md)** objects in the specified chart or chart group.


## Remarks

Use the  **[SeriesCollection](http://msdn.microsoft.com/library/chart-seriescollection-method-powerpoint%28Office.15%29.aspx)** method to return the **SeriesCollection** collection.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

 Use the **[Extend](http://msdn.microsoft.com/library/seriescollection-extend-method-powerpoint%28Office.15%29.aspx)** method to extend an existing series. The following example adds the data in cells C6:C10 in the chart's worksheet to an existing series in the series collection of the chart.




```
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection.Extend "='Sheet1'!$C$6:$C$10"

    End If

End With
```




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

Use the  **[Add](http://msdn.microsoft.com/library/seriescollection-add-method-powerpoint%28Office.15%29.aspx)** method to create a new series and add it to the chart. The following example adds the data from cells D1:D5 in the chart's worksheet as a new series to the chart.




```
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection.Add "='Sheet1'!$D$1:$D$5"

    End If

End With
```




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

Use  **SeriesCollection** ( _Index_ ), where _Index_ is the series index number or name, to return a single **Series** object. The following example sets the color of the interior for the first series in embedded chart one the active document.




```
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection(1).Interior.Color = RGB(255, 0, 0)

    End If

End With
```


## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/seriescollection-add-method-powerpoint%28Office.15%29.aspx)|
|[Extend](http://msdn.microsoft.com/library/seriescollection-extend-method-powerpoint%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/seriescollection-item-method-powerpoint%28Office.15%29.aspx)|
|[NewSeries](http://msdn.microsoft.com/library/seriescollection-newseries-method-powerpoint%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/seriescollection-application-property-powerpoint%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/seriescollection-count-property-powerpoint%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/seriescollection-creator-property-powerpoint%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/seriescollection-parent-property-powerpoint%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/object-model-powerpoint-vba-reference%28Office.15%29.aspx)
