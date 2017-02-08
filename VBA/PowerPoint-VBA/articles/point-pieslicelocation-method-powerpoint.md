---
title: Point.PieSliceLocation Method (PowerPoint)
keywords: vbapp10.chm714011
f1_keywords:
- vbapp10.chm714011
ms.prod: POWERPOINT
ms.assetid: 9af5f72b-3626-9f49-09e5-6fdde51f238e
---


# Point.PieSliceLocation Method (PowerPoint)

Returns the vertical or horizontal position, in points, of a point on a chart item from the top or left edge of the object to the top or left edge of the chart area.


## Syntax

 _expression_. **PieSliceLocation**( **_loc_**, **_Index_** )

 _expression_ A variable that represents a **Point** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _loc_|Required|**[XlPieSliceLocation](http://msdn.microsoft.com/library/xlpieslicelocation-enumeration-excel%28Office.15%29.aspx)**|Specifies a horizontal or vertical coordinate.|
| _Index_|Optional|**[XlPieSliceIndex](http://msdn.microsoft.com/library/xlpiesliceindex-enumeration-excel%28Office.15%29.aspx)**|Specifies which pie slice position coordinate to return. The default is  **xlOuterCenterPoint**.|

### Return Value

Double


## Remarks

This property applies only to pie chart types.


## See also


#### Concepts


[Point Object](point-object-powerpoint.md)

