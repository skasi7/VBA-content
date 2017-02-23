---
title: Axis Members (Excel)
ms.prod: EXCEL
ms.assetid: 2b60f79e-339d-a6cf-7ec6-a915b550c634
---


# Axis Members (Excel)
Represents a single axis in a chart.

Represents a single axis in a chart.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Delete](axis-delete-method-excel.md)|Deletes the object.|
|[Select](axis-select-method-excel.md)|Selects the object.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](axis-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[AxisBetweenCategories](axis-axisbetweencategories-property-excel.md)| **True** if the value axis crosses the category axis between categories. Read/write **Boolean** .|
|[AxisGroup](axis-axisgroup-property-excel.md)|Returns the group for the specified axis. Read-only|
|[AxisTitle](axis-axistitle-property-excel.md)|Returns an  **[AxisTitle](axistitle-object-excel.md)** object that represents the title of the specified axis. Read-only.|
|[BaseUnit](axis-baseunit-property-excel.md)|Returns or sets the base unit for the specified category axis. Read/write  **[XlTimeUnit](xltimeunit-enumeration-excel.md)** .|
|[BaseUnitIsAuto](axis-baseunitisauto-property-excel.md)| **True** if Microsoft Excel chooses appropriate base units for the specified category axis. The default value is **True** . Read/write **Boolean** .|
|[Border](axis-border-property-excel.md)|Returns a  **[Border](border-object-excel.md)** object that represents the border of the object.|
|[CategoryNames](axis-categorynames-property-excel.md)|Returns or sets all the category names for the specified axis, as a text array. When you set this property, you can set it to either an array or a  **[Range](range-object-excel.md)** object that contains the category names. Read/write **Variant** .|
|[CategoryType](axis-categorytype-property-excel.md)|Returns or sets the category axis type. Read/write  **[XlCategoryType](xlcategorytype-enumeration-excel.md)** .|
|[Creator](axis-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[Crosses](axis-crosses-property-excel.md)|Returns or sets the point on the specified axis where the other axis crosses. Read/write  **Long** .|
|[CrossesAt](axis-crossesat-property-excel.md)|Returns or sets the point on the value axis where the category axis crosses it. Applies only to the value axis. Read/write  **Double** .|
|[DisplayUnit](axis-displayunit-property-excel.md)|Returns or sets the unit label for the value axis. Read/write  **[XlDisplayUnit](xldisplayunit-enumeration-excel.md)** , **xlCustom** , or **xlNone** .|
|[DisplayUnitCustom](axis-displayunitcustom-property-excel.md)|If the value of the  **[DisplayUnit](axis-displayunit-property-excel.md)** property is **xlCustom** , the **DisplayUnitCustom** property returns or sets the value of the displayed units. The value must be from 0 through 10E307. Read/write **Double** .|
|[DisplayUnitLabel](axis-displayunitlabel-property-excel.md)|Returns the  **[DisplayUnitLabel](displayunitlabel-object-excel.md)** object for the specified axis. Returns **null** if the **[HasDisplayUnitLabel](axis-hasdisplayunitlabel-property-excel.md)** property is set to **False** . Read-only.|
|[Format](axis-format-property-excel.md)|Returns the  **[ChartFormat](chartformat-object-excel.md)** object. Read-only.|
|[HasDisplayUnitLabel](axis-hasdisplayunitlabel-property-excel.md)| **True** if the label specified by the **[DisplayUnit](axis-displayunit-property-excel.md)** or **[DisplayUnitCustom](axis-displayunitcustom-property-excel.md)** property is displayed on the specified axis. The default value is **True** . Read/write **Boolean** .|
|[HasMajorGridlines](axis-hasmajorgridlines-property-excel.md)| **True** if the axis has major gridlines. Only axes in the primary axis group can have gridlines. Read/write **Boolean** .|
|[HasMinorGridlines](axis-hasminorgridlines-property-excel.md)| **True** if the axis has minor gridlines. Only axes in the primary axis group can have gridlines. Read/write **Boolean** .|
|[HasTitle](axis-hastitle-property-excel.md)| **True** if the axis or chart has a visible title. Read/write **Boolean** .|
|[Height](axis-height-property-excel.md)|Returns a  **Double** value that represents the height, in points, of the object.|
|[Left](axis-left-property-excel.md)|Returns a  **Double** value that represents the distance, in points, from the left edge of the object to the left edge of the chart area.|
|[LogBase](axis-logbase-property-excel.md)|Returns or sets the base of the logarithm when you are using log scales. Read/write  **Double** .|
|[MajorGridlines](axis-majorgridlines-property-excel.md)|Returns a  **[Gridlines](gridlines-object-excel.md)** object that represents the major gridlines for the specified axis. Only axes in the primary axis group can have gridlines. Read-only.|
|[MajorTickMark](axis-majortickmark-property-excel.md)|Returns or sets the type of major tick mark for the specified axis. Read/write  **[XlTickMark](xltickmark-enumeration-excel.md)** .|
|[MajorUnit](axis-majorunit-property-excel.md)|Returns or sets the major units for the value axis. Read/write  **Double** .|
|[MajorUnitIsAuto](axis-majorunitisauto-property-excel.md)| **True** if Microsoft Excel calculates the major units for the value axis. Read/write **Boolean** .|
|[MajorUnitScale](axis-majorunitscale-property-excel.md)|Returns or sets the major unit scale value for the category axis when the  **CategoryType** property is set to **xlTimeScale** . Read/write **[XlTimeUnit](xltimeunit-enumeration-excel.md)** .|
|[MaximumScale](axis-maximumscale-property-excel.md)|Returns or sets the maximum value on the value axis. Read/write  **Double** .|
|[MaximumScaleIsAuto](axis-maximumscaleisauto-property-excel.md)| **True** if Microsoft Excel calculates the maximum value for the value axis. Read/write **Boolean** .|
|[MinimumScale](axis-minimumscale-property-excel.md)|Returns or sets the minimum value on the value axis. Read/write  **Double** .|
|[MinimumScaleIsAuto](axis-minimumscaleisauto-property-excel.md)| **True** if Microsoft Excel calculates the minimum value for the value axis. Read/write **Boolean** .|
|[MinorGridlines](axis-minorgridlines-property-excel.md)|Returns a  **[Gridlines](gridlines-object-excel.md)** object that represents the minor gridlines for the specified axis. Only axes in the primary axis group can have gridlines. Read-only.|
|[MinorTickMark](axis-minortickmark-property-excel.md)|Returns or sets the type of minor tick mark for the specified axis. Read/write  **[XlTickMark](xltickmark-enumeration-excel.md)** .|
|[MinorUnit](axis-minorunit-property-excel.md)|Returns or sets the minor units on the value axis. Read/write  **Double** .|
|[MinorUnitIsAuto](axis-minorunitisauto-property-excel.md)| **True** if Microsoft Excel calculates minor units for the value axis. Read/write **Boolean** .|
|[MinorUnitScale](axis-minorunitscale-property-excel.md)|Returns or sets the minor unit scale value for the category axis when the  **CategoryType** property is set to **xlTimeScale** . Read/write **[XlTimeUnit](xltimeunit-enumeration-excel.md)** .|
|[Parent](axis-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[ReversePlotOrder](axis-reverseplotorder-property-excel.md)| **True** if Microsoft Excel plots data points from last to first. Read/write **Boolean** .|
|[ScaleType](axis-scaletype-property-excel.md)|Returns or sets the value axis scale type. Read/write  **[XlScaleType](xlscaletype-enumeration-excel.md)** .|
|[TickLabelPosition](axis-ticklabelposition-property-excel.md)|Describes the position of tick-mark labels on the specified axis. Read/write  **[XlTickLabelPosition](xlticklabelposition-enumeration-excel.md)** .|
|[TickLabels](axis-ticklabels-property-excel.md)|Returns a  **[TickLabels](ticklabels-object-excel.md)** object that represents the tick-mark labels for the specified axis. Read-only.|
|[TickLabelSpacing](axis-ticklabelspacing-property-excel.md)|Returns or sets the number of categories or series between tick-mark labels. Applies only to category and series axes. Can be a value from 1 through 31999. Read/write  **Long** .|
|[TickLabelSpacingIsAuto](axis-ticklabelspacingisauto-property-excel.md)|Returns or sets whether or not the tick label spacing is automatic. Read/write  **Boolean** .|
|[TickMarkSpacing](axis-tickmarkspacing-property-excel.md)|Returns or sets the number of categories or series between tick marks. Applies only to category and series axes. Can be a value from 1 through 31999. Read/write  **Long** .|
|[Top](axis-top-property-excel.md)|Returns a  **Double** value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).|
|[Type](axis-type-property-excel.md)|Returns an  **[XlAxisType](xlaxistype-enumeration-excel.md)** value that represents the Axis type.|
|[Width](axis-width-property-excel.md)|Returns a  **Double** value that represents the width, in points, of the object.|

