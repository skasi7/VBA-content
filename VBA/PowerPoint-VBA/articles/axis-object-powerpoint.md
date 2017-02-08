---
title: Axis Object (PowerPoint)
keywords: vbapp10.chm682000
f1_keywords:
- vbapp10.chm682000
ms.prod: POWERPOINT
ms.assetid: 38d5e006-ac32-7bdb-f9f0-e8a858dcbf49
---


# Axis Object (PowerPoint)

Represents a single axis in a chart.


## Remarks

The  **Axis** object is a member of the **[Axes](http://msdn.microsoft.com/library/axes-object-powerpoint%28Office.15%29.aspx)** collection.

Use  **Axes** ( _Type_, _AxisGroup_ ) where _Type_ is the axis type and _AxisGroup_ is the axis group to return a single **Axis** object. _Type_ can be one of the following **[XlAxisType](http://msdn.microsoft.com/library/xlaxistype-enumeration-powerpoint%28Office.15%29.aspx)** constants: **xlCategory**, **xlSeries**, or **xlValue**. _AxisGroup_ can be one of the following **[XlAxisGroup](http://msdn.microsoft.com/library/xlaxisgroup-enumeration-powerpoint%28Office.15%29.aspx)** constants: **xlPrimary** or **xlSecondary**. For more information, see the **[Axes](http://msdn.microsoft.com/library/chart-axes-method-powerpoint%28Office.15%29.aspx)** method.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the category axis title text for the first chart in the active document.




```
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.Axes(xlCategory)

            .HasTitle = True

            .AxisTitle.Caption = "1994"

        End With

    End If

End With
```


## Methods



|**Name**|
|:-----|
|[Delete](http://msdn.microsoft.com/library/axis-delete-method-powerpoint%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/axis-select-method-powerpoint%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/axis-application-property-powerpoint%28Office.15%29.aspx)|
|[AxisBetweenCategories](http://msdn.microsoft.com/library/axis-axisbetweencategories-property-powerpoint%28Office.15%29.aspx)|
|[AxisGroup](http://msdn.microsoft.com/library/axis-axisgroup-property-powerpoint%28Office.15%29.aspx)|
|[AxisTitle](http://msdn.microsoft.com/library/axis-axistitle-property-powerpoint%28Office.15%29.aspx)|
|[BaseUnit](http://msdn.microsoft.com/library/axis-baseunit-property-powerpoint%28Office.15%29.aspx)|
|[BaseUnitIsAuto](http://msdn.microsoft.com/library/axis-baseunitisauto-property-powerpoint%28Office.15%29.aspx)|
|[Border](http://msdn.microsoft.com/library/axis-border-property-powerpoint%28Office.15%29.aspx)|
|[CategoryNames](http://msdn.microsoft.com/library/axis-categorynames-property-powerpoint%28Office.15%29.aspx)|
|[CategoryType](http://msdn.microsoft.com/library/axis-categorytype-property-powerpoint%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/axis-creator-property-powerpoint%28Office.15%29.aspx)|
|[Crosses](http://msdn.microsoft.com/library/axis-crosses-property-powerpoint%28Office.15%29.aspx)|
|[CrossesAt](http://msdn.microsoft.com/library/axis-crossesat-property-powerpoint%28Office.15%29.aspx)|
|[DisplayUnit](http://msdn.microsoft.com/library/axis-displayunit-property-powerpoint%28Office.15%29.aspx)|
|[DisplayUnitCustom](http://msdn.microsoft.com/library/axis-displayunitcustom-property-powerpoint%28Office.15%29.aspx)|
|[DisplayUnitLabel](http://msdn.microsoft.com/library/axis-displayunitlabel-property-powerpoint%28Office.15%29.aspx)|
|[Format](http://msdn.microsoft.com/library/axis-format-property-powerpoint%28Office.15%29.aspx)|
|[HasDisplayUnitLabel](http://msdn.microsoft.com/library/axis-hasdisplayunitlabel-property-powerpoint%28Office.15%29.aspx)|
|[HasMajorGridlines](http://msdn.microsoft.com/library/axis-hasmajorgridlines-property-powerpoint%28Office.15%29.aspx)|
|[HasMinorGridlines](http://msdn.microsoft.com/library/axis-hasminorgridlines-property-powerpoint%28Office.15%29.aspx)|
|[HasTitle](http://msdn.microsoft.com/library/axis-hastitle-property-powerpoint%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/axis-height-property-powerpoint%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/axis-left-property-powerpoint%28Office.15%29.aspx)|
|[LogBase](http://msdn.microsoft.com/library/axis-logbase-property-powerpoint%28Office.15%29.aspx)|
|[MajorGridlines](http://msdn.microsoft.com/library/axis-majorgridlines-property-powerpoint%28Office.15%29.aspx)|
|[MajorTickMark](http://msdn.microsoft.com/library/axis-majortickmark-property-powerpoint%28Office.15%29.aspx)|
|[MajorUnit](http://msdn.microsoft.com/library/axis-majorunit-property-powerpoint%28Office.15%29.aspx)|
|[MajorUnitIsAuto](http://msdn.microsoft.com/library/axis-majorunitisauto-property-powerpoint%28Office.15%29.aspx)|
|[MajorUnitScale](http://msdn.microsoft.com/library/axis-majorunitscale-property-powerpoint%28Office.15%29.aspx)|
|[MaximumScale](http://msdn.microsoft.com/library/axis-maximumscale-property-powerpoint%28Office.15%29.aspx)|
|[MaximumScaleIsAuto](http://msdn.microsoft.com/library/axis-maximumscaleisauto-property-powerpoint%28Office.15%29.aspx)|
|[MinimumScale](http://msdn.microsoft.com/library/axis-minimumscale-property-powerpoint%28Office.15%29.aspx)|
|[MinimumScaleIsAuto](http://msdn.microsoft.com/library/axis-minimumscaleisauto-property-powerpoint%28Office.15%29.aspx)|
|[MinorGridlines](http://msdn.microsoft.com/library/axis-minorgridlines-property-powerpoint%28Office.15%29.aspx)|
|[MinorTickMark](http://msdn.microsoft.com/library/axis-minortickmark-property-powerpoint%28Office.15%29.aspx)|
|[MinorUnit](http://msdn.microsoft.com/library/axis-minorunit-property-powerpoint%28Office.15%29.aspx)|
|[MinorUnitIsAuto](http://msdn.microsoft.com/library/axis-minorunitisauto-property-powerpoint%28Office.15%29.aspx)|
|[MinorUnitScale](http://msdn.microsoft.com/library/axis-minorunitscale-property-powerpoint%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/axis-parent-property-powerpoint%28Office.15%29.aspx)|
|[ReversePlotOrder](http://msdn.microsoft.com/library/axis-reverseplotorder-property-powerpoint%28Office.15%29.aspx)|
|[ScaleType](http://msdn.microsoft.com/library/axis-scaletype-property-powerpoint%28Office.15%29.aspx)|
|[TickLabelPosition](http://msdn.microsoft.com/library/axis-ticklabelposition-property-powerpoint%28Office.15%29.aspx)|
|[TickLabels](http://msdn.microsoft.com/library/axis-ticklabels-property-powerpoint%28Office.15%29.aspx)|
|[TickLabelSpacing](http://msdn.microsoft.com/library/axis-ticklabelspacing-property-powerpoint%28Office.15%29.aspx)|
|[TickLabelSpacingIsAuto](http://msdn.microsoft.com/library/axis-ticklabelspacingisauto-property-powerpoint%28Office.15%29.aspx)|
|[TickMarkSpacing](http://msdn.microsoft.com/library/axis-tickmarkspacing-property-powerpoint%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/axis-top-property-powerpoint%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/axis-type-property-powerpoint%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/axis-width-property-powerpoint%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/object-model-powerpoint-vba-reference%28Office.15%29.aspx)
