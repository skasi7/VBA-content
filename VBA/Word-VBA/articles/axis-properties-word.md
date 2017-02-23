---
title: Axis Properties (Word)
ms.prod: WORD
ms.assetid: aa91fac2-8d22-42a3-8f4e-978337158c3f
---


# Axis Properties (Word)

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](axis-application-property-word.md)|When used without an object qualifier, returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application. When used with an object qualifier, returns an **Application** object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.|
|[AxisBetweenCategories](axis-axisbetweencategories-property-word.md)| **True** if the value axis crosses the category axis between categories. Read/write **Boolean** .|
|[AxisGroup](axis-axisgroup-property-word.md)|Returns the type of axis group. Read-only  **[XlAxisGroup](xlaxisgroup-enumeration-word.md)** .|
|[AxisTitle](axis-axistitle-property-word.md)|Returns the title of the specified axis. Read-only  **[AxisTitle](axistitle-object-word.md)** .|
|[BaseUnit](axis-baseunit-property-word.md)|Returns or sets the base unit for the specified category axis. Read/write  **[XlTimeUnit](xltimeunit-enumeration-word.md)** .|
|[BaseUnitIsAuto](axis-baseunitisauto-property-word.md)| **True** if Microsoft Word chooses appropriate base units for the specified category axis. The default is **True** . Read/write **Boolean** .|
|[Border](axis-border-property-word.md)|Returns the border of the object. Read-only  **[ChartBorder](chartborder-object-word.md)** .|
|[CategoryNames](axis-categorynames-property-word.md)|Returns or sets all the category names as a text array for the specified axis. Read/write  **Variant** .|
|[CategoryType](axis-categorytype-property-word.md)|Returns or sets the category axis type. Read/write  **[XlCategoryType](xlcategorytype-enumeration-word.md)** .|
|[Creator](axis-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[Crosses](axis-crosses-property-word.md)|Returns or sets the point on the specified axis where the other axis crosses. Read/write  **Long** .|
|[CrossesAt](axis-crossesat-property-word.md)|Returns or sets the point on the value axis where the category axis crosses it. Applies only to the value axis. Read/write  **Double** .|
|[DisplayUnit](axis-displayunit-property-word.md)|Returns or sets the unit label for the value axis. Read/write  **[XlDisplayUnit](xldisplayunit-enumeration-word.md)** , **xlCustom** , or **xlNone** .|
|[DisplayUnitCustom](axis-displayunitcustom-property-word.md)|If the value of the  **[DisplayUnit](axis-displayunit-property-word.md)** property is **xlCustom** , returns or sets the value of the displayed units. Read/write **Double** .|
|[DisplayUnitLabel](axis-displayunitlabel-property-word.md)|Returns the  **[DisplayUnitLabel](displayunitlabel-object-word.md)** object for the specified axis. Returns **null** if the **[HasDisplayUnitLabel](axis-hasdisplayunitlabel-property-word.md)** property is set to **False** . Read-only.|
|[Format](axis-format-property-word.md)|Returns the line, fill, and effect formatting for the object. Read-only  **[ChartFormat](chartformat-object-word.md)** .|
|[HasDisplayUnitLabel](axis-hasdisplayunitlabel-property-word.md)| **True** if the label specified by the **[DisplayUnit](axis-displayunit-property-word.md)** or **[DisplayUnitCustom](axis-displayunitcustom-property-word.md)** property is displayed on the specified axis. The default is **True** . Read/write **Boolean** .|
|[HasMajorGridlines](axis-hasmajorgridlines-property-word.md)| **True** if the axis has major gridlines. Read/write **Boolean** .|
|[HasMinorGridlines](axis-hasminorgridlines-property-word.md)| **True** if the axis has minor gridlines. Read/write **Boolean** .|
|[HasTitle](axis-hastitle-property-word.md)| **True** if the axis or chart has a visible title. Read/write **Boolean** .|
|[Height](axis-height-property-word.md)|Returns the height, in points, of the object. Read-only  **Double** .|
|[Left](axis-left-property-word.md)|Returns the distance, in points, from the left edge of the object to the left edge of the chart area. Read-only  **Double** .|
|[LogBase](axis-logbase-property-word.md)|Returns or sets the base of the logarithm when you are using log scales. Read/write  **Double** .|
|[MajorGridlines](axis-majorgridlines-property-word.md)|Returns the major gridlines for the specified axis. Read-only  **[Gridlines](gridlines-object-word.md)** .|
|[MajorTickMark](axis-majortickmark-property-word.md)|Returns or sets the type of major tick mark for the specified axis. Read/write  **[XlTickMark](xltickmark-enumeration-word.md)** .|
|[MajorUnit](axis-majorunit-property-word.md)|Returns or sets the major units for the value axis. Read/write  **Double** .|
|[MajorUnitIsAuto](axis-majorunitisauto-property-word.md)| **True** if Microsoft Word calculates the major units for the value axis. Read/write **Boolean** .|
|[MajorUnitScale](axis-majorunitscale-property-word.md)|Returns or sets the major unit scale value for the category axis when the  **[CategoryType](axis-categorytype-property-word.md)** property is set to **xlTimeScale** . Read/write **[XlTimeUnit](xltimeunit-enumeration-word.md)** .|
|[MaximumScale](axis-maximumscale-property-word.md)|Returns or sets the maximum value on the value axis. Read/write  **Double** .|
|[MaximumScaleIsAuto](axis-maximumscaleisauto-property-word.md)| **True** if Microsoft Word calculates the maximum value for the value axis. Read/write **Boolean** .|
|[MinimumScale](axis-minimumscale-property-word.md)|Returns or sets the minimum value on the value axis. Read/write  **Double** .|
|[MinimumScaleIsAuto](axis-minimumscaleisauto-property-word.md)| **True** if Microsoft Word calculates the minimum value for the value axis. Read/write **Boolean** .|
|[MinorGridlines](axis-minorgridlines-property-word.md)|Returns the minor gridlines for the specified axis. Read-only  **[Gridlines](gridlines-object-word.md)** .|
|[MinorTickMark](axis-minortickmark-property-word.md)|Returns or sets the type of minor tick mark for the specified axis. Read/write  **[XlTickMark](xltickmark-enumeration-word.md)** .|
|[MinorUnit](axis-minorunit-property-word.md)|Returns or sets the minor units on the value axis. Read/write  **Double** .|
|[MinorUnitIsAuto](axis-minorunitisauto-property-word.md)| **True** if Microsoft Word calculates minor units for the value axis. Read/write **Boolean** .|
|[MinorUnitScale](axis-minorunitscale-property-word.md)|Returns or sets the minor unit scale value for the category axis when the  **[CategoryType](axis-categorytype-property-word.md)** property is set to **xlTimeScale** . Read/write **[XlTimeUnit](xltimeunit-enumeration-word.md)** .|
|[Parent](axis-parent-property-word.md)|Returns the parent for the specified object. Read-only  **Object** .|
|[ReversePlotOrder](axis-reverseplotorder-property-word.md)| **True** if Microsoft Word plots data points from last to first. Read/write **Boolean** .|
|[ScaleType](axis-scaletype-property-word.md)|Returns or sets the value axis scale type. Read/write  **[XlScaleType](xlscaletype-enumeration-word.md)** .|
|[TickLabelPosition](axis-ticklabelposition-property-word.md)|Describes the position of tick-mark labels on the specified axis. Read/write  **[XlTickLabelPosition](xlticklabelposition-enumeration-word.md)** .|
|[TickLabels](axis-ticklabels-property-word.md)|Returns the tick-mark labels for the specified axis. Read-only  **[TickLabels](ticklabels-object-word.md)** .|
|[TickLabelSpacing](axis-ticklabelspacing-property-word.md)|Returns or sets the number of categories or series between tick-mark labels. Read/write  **Long** .|
|[TickLabelSpacingIsAuto](axis-ticklabelspacingisauto-property-word.md)|Returns or sets a value that indicates whether the tick label spacing is automatic. Read/write  **Boolean** .|
|[TickMarkSpacing](axis-tickmarkspacing-property-word.md)|Returns or sets the number of categories or series between tick marks. Read/write  **Long** .|
|[Top](axis-top-property-word.md)|Returns the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart). Read-only  **Double** .|
|[Type](axis-type-property-word.md)|Returns the axis type. Read-only  **[XlAxisType](xlaxistype-enumeration-word.md)** .|
|[Width](axis-width-property-word.md)|Returns the width, in points, of the object. Read-only  **Double** .|

