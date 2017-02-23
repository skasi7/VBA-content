---
title: Axis Properties (PowerPoint)
ms.prod: POWERPOINT
ms.assetid: 562be71b-a2b8-489a-bedc-94441154773c
---


# Axis Properties (PowerPoint)

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](axis-application-property-powerpoint.md)|When used without an object qualifier, returns an  **[Application](application-object-powerpoint.md)** object that represents the Microsoft PowerPoint application. When used with an object qualifier, returns an **Application** object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.|
|[AxisBetweenCategories](axis-axisbetweencategories-property-powerpoint.md)|**True** if the value axis crosses the category axis between categories. Read/write **Boolean**.|
|[AxisGroup](axis-axisgroup-property-powerpoint.md)|Returns the type of axis group. Read-only  **[XlAxisGroup](xlaxisgroup-enumeration-powerpoint.md)**.|
|[AxisTitle](axis-axistitle-property-powerpoint.md)|Returns the title of the specified axis. Read-only  **[AxisTitle](axistitle-object-powerpoint.md)**.|
|[BaseUnit](axis-baseunit-property-powerpoint.md)|Returns or sets the base unit for the specified category axis. Read/write  **[XlTimeUnit](xltimeunit-enumeration-powerpoint.md)**.|
|[BaseUnitIsAuto](axis-baseunitisauto-property-powerpoint.md)|**True** if Microsoft Word chooses appropriate base units for the specified category axis. The default is **True**. Read/write **Boolean**.|
|[Border](axis-border-property-powerpoint.md)|Returns the border of the object. Read-only  **[ChartBorder](chartborder-object-powerpoint.md)**.|
|[CategoryNames](axis-categorynames-property-powerpoint.md)|Returns or sets all the category names as a text array for the specified axis. Read/write  **Variant**.|
|[CategoryType](axis-categorytype-property-powerpoint.md)|Returns or sets the category axis type. Read/write  **[XlCategoryType](xlcategorytype-enumeration-powerpoint.md)**.|
|[Creator](axis-creator-property-powerpoint.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.|
|[Crosses](axis-crosses-property-powerpoint.md)|Returns or sets the point on the specified axis where the other axis crosses. Read/write  **Long**.|
|[CrossesAt](axis-crossesat-property-powerpoint.md)|Returns or sets the point on the value axis where the category axis crosses it. Applies only to the value axis. Read/write  **Double**.|
|[DisplayUnit](axis-displayunit-property-powerpoint.md)|Returns or sets the unit label for the value axis. Read/write  **[XlDisplayUnit](xldisplayunit-enumeration-powerpoint.md)**, **xlCustom**, or **xlNone**.|
|[DisplayUnitCustom](axis-displayunitcustom-property-powerpoint.md)|If the value of the  **[DisplayUnit](axis-displayunit-property-powerpoint.md)** property is **xlCustom**, returns or sets the value of the displayed units. Read/write **Double**.|
|[DisplayUnitLabel](axis-displayunitlabel-property-powerpoint.md)|Returns the  **[DisplayUnitLabel](displayunitlabel-object-powerpoint.md)** object for the specified axis. Returns **null** if the **[HasDisplayUnitLabel](axis-hasdisplayunitlabel-property-powerpoint.md)** property is set to **False**. Read-only.|
|[Format](axis-format-property-powerpoint.md)|Returns the line, fill, and effect formatting for the object. Read-only  **[ChartFormat](chartformat-object-powerpoint.md)**.|
|[HasDisplayUnitLabel](axis-hasdisplayunitlabel-property-powerpoint.md)|**True** if the label specified by the **[DisplayUnit](axis-displayunit-property-powerpoint.md)** or **[DisplayUnitCustom](axis-displayunitcustom-property-powerpoint.md)** property is displayed on the specified axis. The default is **True**. Read/write **Boolean**.|
|[HasMajorGridlines](axis-hasmajorgridlines-property-powerpoint.md)|**True** if the axis has major gridlines. Read/write **Boolean**.|
|[HasMinorGridlines](axis-hasminorgridlines-property-powerpoint.md)|**True** if the axis has minor gridlines. Read/write **Boolean**.|
|[HasTitle](axis-hastitle-property-powerpoint.md)|**True** if the axis or chart has a visible title. Read/write **Boolean**.|
|[Height](axis-height-property-powerpoint.md)|Returns the height, in points, of the object. Read-only  **Double**.|
|[Left](axis-left-property-powerpoint.md)|Returns the distance, in points, from the left edge of the object to the left edge of the chart area. Read-only  **Double**.|
|[LogBase](axis-logbase-property-powerpoint.md)|Returns or sets the base of the logarithm when you are using log scales. Read/write  **Double**.|
|[MajorGridlines](axis-majorgridlines-property-powerpoint.md)|Returns the major gridlines for the specified axis. Read-only  **[Gridlines](gridlines-object-powerpoint.md)**.|
|[MajorTickMark](axis-majortickmark-property-powerpoint.md)|Returns or sets the type of major tick mark for the specified axis. Read/write  **[XlTickMark](xltickmark-enumeration-powerpoint.md)**.|
|[MajorUnit](axis-majorunit-property-powerpoint.md)|Returns or sets the major units for the value axis. Read/write  **Double**.|
|[MajorUnitIsAuto](axis-majorunitisauto-property-powerpoint.md)|**True** if Microsoft Word calculates the major units for the value axis. Read/write **Boolean**.|
|[MajorUnitScale](axis-majorunitscale-property-powerpoint.md)|Returns or sets the major unit scale value for the category axis when the  **[CategoryType](axis-categorytype-property-powerpoint.md)** property is set to **xlTimeScale**. Read/write **[XlTimeUnit](xltimeunit-enumeration-powerpoint.md)**.|
|[MaximumScale](axis-maximumscale-property-powerpoint.md)|Returns or sets the maximum value on the value axis. Read/write  **Double**.|
|[MaximumScaleIsAuto](axis-maximumscaleisauto-property-powerpoint.md)|**True** if Microsoft Word calculates the maximum value for the value axis. Read/write **Boolean**.|
|[MinimumScale](axis-minimumscale-property-powerpoint.md)|Returns or sets the minimum value on the value axis. Read/write  **Double**.|
|[MinimumScaleIsAuto](axis-minimumscaleisauto-property-powerpoint.md)|**True** if Microsoft Word calculates the minimum value for the value axis. Read/write **Boolean**.|
|[MinorGridlines](axis-minorgridlines-property-powerpoint.md)|Returns the minor gridlines for the specified axis. Read-only  **[Gridlines](gridlines-object-powerpoint.md)**.|
|[MinorTickMark](axis-minortickmark-property-powerpoint.md)|Returns or sets the type of minor tick mark for the specified axis. Read/write  **[XlTickMark](xltickmark-enumeration-powerpoint.md)**.|
|[MinorUnit](axis-minorunit-property-powerpoint.md)|Returns or sets the minor units on the value axis. Read/write  **Double**.|
|[MinorUnitIsAuto](axis-minorunitisauto-property-powerpoint.md)|**True** if Microsoft Word calculates minor units for the value axis. Read/write **Boolean**.|
|[MinorUnitScale](axis-minorunitscale-property-powerpoint.md)|Returns or sets the minor unit scale value for the category axis when the  **[CategoryType](axis-categorytype-property-powerpoint.md)** property is set to **xlTimeScale**. Read/write **[XlTimeUnit](xltimeunit-enumeration-powerpoint.md)**.|
|[Parent](axis-parent-property-powerpoint.md)|Returns the parent for the specified object. Read-only  **Object**.|
|[ReversePlotOrder](axis-reverseplotorder-property-powerpoint.md)|**True** if Microsoft Word plots data points from last to first. Read/write **Boolean**.|
|[ScaleType](axis-scaletype-property-powerpoint.md)|Returns or sets the value axis scale type. Read/write  **[XlScaleType](xlscaletype-enumeration-powerpoint.md)**.|
|[TickLabelPosition](axis-ticklabelposition-property-powerpoint.md)|Describes the position of tick-mark labels on the specified axis. Read/write  **[XlTickLabelPosition](xlticklabelposition-enumeration-powerpoint.md)**.|
|[TickLabels](axis-ticklabels-property-powerpoint.md)|Returns the tick-mark labels for the specified axis. Read-only  **[TickLabels](ticklabels-object-powerpoint.md)**.|
|[TickLabelSpacing](axis-ticklabelspacing-property-powerpoint.md)|Returns or sets the number of categories or series between tick-mark labels. Read/write  **Long**.|
|[TickLabelSpacingIsAuto](axis-ticklabelspacingisauto-property-powerpoint.md)|Returns or sets a value that indicates whether the tick label spacing is automatic. Read/write  **Boolean**.|
|[TickMarkSpacing](axis-tickmarkspacing-property-powerpoint.md)|Returns or sets the number of categories or series between tick marks. Read/write  **Long**.|
|[Top](axis-top-property-powerpoint.md)|Returns the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart). Read-only  **Double**.|
|[Type](axis-type-property-powerpoint.md)|Returns the axis type. Read-only  **[XlAxisType](xlaxistype-enumeration-powerpoint.md)**.|
|[Width](axis-width-property-powerpoint.md)|Returns the width, in points, of the object. Read-only  **Double**.|

