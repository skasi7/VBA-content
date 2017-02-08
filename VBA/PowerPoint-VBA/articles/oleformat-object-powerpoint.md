---
title: OLEFormat Object (PowerPoint)
keywords: vbapp10.chm562000
f1_keywords:
- vbapp10.chm562000
ms.prod: POWERPOINT
ms.assetid: fbb6d6dd-4dbb-461b-986e-5095c6dc1486
---


# OLEFormat Object (PowerPoint)

Contains properties and methods that apply to OLE objects. 


## Remarks

The  **[LinkFormat](linkformat-object-powerpoint.md)** object contains properties and methods that apply to linked OLE objects only. The **[PictureFormat](pictureformat-object-powerpoint.md)** object contains properties and methods that apply to pictures and OLE objects.


## Example

Use the  **OLEFormat** property to return an **OLEFormat** object. The following example loops through all the shapes on all the slides in the active presentation and sets all linked Microsoft Excel worksheets to be updated manually.


```
For Each sld In ActivePresentation.Slides

    For Each sh In sld.Shapes

        If sh.Type = msoLinkedOLEObject Then

            If sh.OLEFormat.ProgID = "Excel.Sheet" Then

                sh.LinkFormat.AutoUpdate = ppUpdateOptionManual

            End If

        End If

    Next

Next
```


## Methods



|**Name**|
|:-----|
|[Activate](http://msdn.microsoft.com/library/oleformat-activate-method-powerpoint%28Office.15%29.aspx)|
|[DoVerb](http://msdn.microsoft.com/library/oleformat-doverb-method-powerpoint%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/oleformat-application-property-powerpoint%28Office.15%29.aspx)|
|[FollowColors](http://msdn.microsoft.com/library/oleformat-followcolors-property-powerpoint%28Office.15%29.aspx)|
|[Object](http://msdn.microsoft.com/library/oleformat-object-property-powerpoint%28Office.15%29.aspx)|
|[ObjectVerbs](http://msdn.microsoft.com/library/oleformat-objectverbs-property-powerpoint%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/oleformat-parent-property-powerpoint%28Office.15%29.aspx)|
|[ProgID](http://msdn.microsoft.com/library/oleformat-progid-property-powerpoint%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/object-model-powerpoint-vba-reference%28Office.15%29.aspx)
