---
title: ShapeRange.Align Method (PowerPoint)
keywords: vbapp10.chm548063
f1_keywords:
- vbapp10.chm548063
ms.prod: POWERPOINT
ms.assetid: 5d4553ad-521a-1f3c-77ba-3dd5fbd02a09
---


# ShapeRange.Align Method (PowerPoint)

Aligns the shapes in the specified range of shapes.


## Syntax

 _expression_. **Align**( **_AlignCmd_**, **_RelativeTo_** )

 _expression_ A variable that represents a **ShapeRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _AlignCmd_|Required|**[MsoAlignCmd](http://msdn.microsoft.com/library/msoaligncmd-enumeration-office%28Office.15%29.aspx)**|Specifies the way the shapes in the specified shape range are to be aligned.|
| _RelativeTo_|Required|**[MsoTriState](http://msdn.microsoft.com/library/msotristate-enumeration-office%28Office.15%29.aspx)**|Determines whether shapes are aligned relative to the edge of the slide.|

## Example

This example aligns the left edges of all the shapes in the specified range in  `myDocument` with the left edge of the leftmost shape in the range.


```vb
Set myDocument = ActivePresentation.Slides(1) 
myDocument.Shapes.Range.Align msoAlignLefts, msoFalse
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-powerpoint.md)

