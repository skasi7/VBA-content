---
title: Series Object (PowerPoint)
keywords: vbapp10.chm716000
f1_keywords:
- vbapp10.chm716000
ms.prod: POWERPOINT
ms.assetid: 5c8c2d92-d8ca-4d21-e213-c374292275d4
---


# Series Object (PowerPoint)

Represents a series in a chart.


## Remarks

 The **Series** object is a member of the **[SeriesCollection](seriescollection-object-powerpoint.md)** collection.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

Use  **[SeriesCollection](http://msdn.microsoft.com/library/chart-seriescollection-method-powerpoint%28Office.15%29.aspx)** ( _Index_ ), where _Index_ is the series index number or name, to return a single **Series** object. The following example sets the color of the interior for the first series of the first chart in the active document.

The series index number indicates the order in which the series were added to the chart.  `SeriesCollection(1)` is the first series added to the chart, and `SeriesCollection(SeriesCollection.Count)` is the last one added.




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
|[ApplyDataLabels](http://msdn.microsoft.com/library/series-applydatalabels-method-powerpoint%28Office.15%29.aspx)|
|[ClearFormats](http://msdn.microsoft.com/library/series-clearformats-method-powerpoint%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/series-copy-method-powerpoint%28Office.15%29.aspx)|
|[DataLabels](http://msdn.microsoft.com/library/series-datalabels-method-powerpoint%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/series-delete-method-powerpoint%28Office.15%29.aspx)|
|[ErrorBar](http://msdn.microsoft.com/library/series-errorbar-method-powerpoint%28Office.15%29.aspx)|
|[Paste](http://msdn.microsoft.com/library/series-paste-method-powerpoint%28Office.15%29.aspx)|
|[Points](http://msdn.microsoft.com/library/series-points-method-powerpoint%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/series-select-method-powerpoint%28Office.15%29.aspx)|
|[Trendlines](http://msdn.microsoft.com/library/series-trendlines-method-powerpoint%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/series-application-property-powerpoint%28Office.15%29.aspx)|
|[ApplyPictToEnd](http://msdn.microsoft.com/library/series-applypicttoend-property-powerpoint%28Office.15%29.aspx)|
|[ApplyPictToFront](http://msdn.microsoft.com/library/series-applypicttofront-property-powerpoint%28Office.15%29.aspx)|
|[ApplyPictToSides](http://msdn.microsoft.com/library/series-applypicttosides-property-powerpoint%28Office.15%29.aspx)|
|[AxisGroup](http://msdn.microsoft.com/library/series-axisgroup-property-powerpoint%28Office.15%29.aspx)|
|[BarShape](http://msdn.microsoft.com/library/series-barshape-property-powerpoint%28Office.15%29.aspx)|
|[BubbleSizes](http://msdn.microsoft.com/library/series-bubblesizes-property-powerpoint%28Office.15%29.aspx)|
|[ChartType](http://msdn.microsoft.com/library/series-charttype-property-powerpoint%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/series-creator-property-powerpoint%28Office.15%29.aspx)|
|[ErrorBars](http://msdn.microsoft.com/library/series-errorbars-property-powerpoint%28Office.15%29.aspx)|
|[Explosion](http://msdn.microsoft.com/library/series-explosion-property-powerpoint%28Office.15%29.aspx)|
|[Format](http://msdn.microsoft.com/library/series-format-property-powerpoint%28Office.15%29.aspx)|
|[Formula](http://msdn.microsoft.com/library/series-formula-property-powerpoint%28Office.15%29.aspx)|
|[FormulaLocal](http://msdn.microsoft.com/library/series-formulalocal-property-powerpoint%28Office.15%29.aspx)|
|[FormulaR1C1](http://msdn.microsoft.com/library/series-formular1c1-property-powerpoint%28Office.15%29.aspx)|
|[FormulaR1C1Local](http://msdn.microsoft.com/library/series-formular1c1local-property-powerpoint%28Office.15%29.aspx)|
|[Has3DEffect](http://msdn.microsoft.com/library/series-has3deffect-property-powerpoint%28Office.15%29.aspx)|
|[HasDataLabels](http://msdn.microsoft.com/library/series-hasdatalabels-property-powerpoint%28Office.15%29.aspx)|
|[HasErrorBars](http://msdn.microsoft.com/library/series-haserrorbars-property-powerpoint%28Office.15%29.aspx)|
|[HasLeaderLines](http://msdn.microsoft.com/library/series-hasleaderlines-property-powerpoint%28Office.15%29.aspx)|
|[InvertColor](http://msdn.microsoft.com/library/series-invertcolor-property-powerpoint%28Office.15%29.aspx)|
|[InvertColorIndex](http://msdn.microsoft.com/library/series-invertcolorindex-property-powerpoint%28Office.15%29.aspx)|
|[InvertIfNegative](http://msdn.microsoft.com/library/series-invertifnegative-property-powerpoint%28Office.15%29.aspx)|
|[IsFiltered](http://msdn.microsoft.com/library/series-isfiltered-property-powerpoint%28Office.15%29.aspx)|
|[LeaderLines](http://msdn.microsoft.com/library/series-leaderlines-property-powerpoint%28Office.15%29.aspx)|
|[MarkerBackgroundColor](http://msdn.microsoft.com/library/series-markerbackgroundcolor-property-powerpoint%28Office.15%29.aspx)|
|[MarkerBackgroundColorIndex](http://msdn.microsoft.com/library/series-markerbackgroundcolorindex-property-powerpoint%28Office.15%29.aspx)|
|[MarkerForegroundColor](http://msdn.microsoft.com/library/series-markerforegroundcolor-property-powerpoint%28Office.15%29.aspx)|
|[MarkerForegroundColorIndex](http://msdn.microsoft.com/library/series-markerforegroundcolorindex-property-powerpoint%28Office.15%29.aspx)|
|[MarkerSize](http://msdn.microsoft.com/library/series-markersize-property-powerpoint%28Office.15%29.aspx)|
|[MarkerStyle](http://msdn.microsoft.com/library/series-markerstyle-property-powerpoint%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/series-name-property-powerpoint%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/series-parent-property-powerpoint%28Office.15%29.aspx)|
|**[ParentDataLabelOption](http://msdn.microsoft.com/library/series-parentdatalabeloption-property-powerpoint%28Office.15%29.aspx)**|
|:-----|
|[PictureType](http://msdn.microsoft.com/library/series-picturetype-property-powerpoint%28Office.15%29.aspx)|
|[PictureUnit2](http://msdn.microsoft.com/library/series-pictureunit2-property-powerpoint%28Office.15%29.aspx)|
|[PlotColorIndex](http://msdn.microsoft.com/library/series-plotcolorindex-property-powerpoint%28Office.15%29.aspx)|
|[PlotOrder](http://msdn.microsoft.com/library/series-plotorder-property-powerpoint%28Office.15%29.aspx)|
|[QuartileCalculationInclusiveMedian](http://msdn.microsoft.com/library/series-quartilecalculationinclusivemedian-property-powerpoint%28Office.15%29.aspx)|
|[Shadow](http://msdn.microsoft.com/library/series-shadow-property-powerpoint%28Office.15%29.aspx)|
|[Smooth](http://msdn.microsoft.com/library/series-smooth-property-powerpoint%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/series-type-property-powerpoint%28Office.15%29.aspx)|
|[Values](http://msdn.microsoft.com/library/series-values-property-powerpoint%28Office.15%29.aspx)|
|[XValues](http://msdn.microsoft.com/library/series-xvalues-property-powerpoint%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/object-model-powerpoint-vba-reference%28Office.15%29.aspx)
