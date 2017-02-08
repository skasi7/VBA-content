---
title: LinkFormat Object (PowerPoint)
keywords: vbapp10.chm563000
f1_keywords:
- vbapp10.chm563000
ms.prod: POWERPOINT
ms.assetid: e89ee344-4197-ac0d-dd53-966e4672a3ce
---


# LinkFormat Object (PowerPoint)

Contains properties and methods that apply to linked OLE objects, linked pictures, and IIRC media objects. 


## Example

Use the  **LinkFormat** property to return a **LinkFormat** object. The following example loops through all the shapes on all the slides in the active presentation and sets all linked Microsoft Excel worksheets to be updated manually.


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
|[BreakLink](http://msdn.microsoft.com/library/linkformat-breaklink-method-powerpoint%28Office.15%29.aspx)|
|[Update](http://msdn.microsoft.com/library/linkformat-update-method-powerpoint%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/linkformat-application-property-powerpoint%28Office.15%29.aspx)|
|[AutoUpdate](http://msdn.microsoft.com/library/linkformat-autoupdate-property-powerpoint%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/linkformat-parent-property-powerpoint%28Office.15%29.aspx)|
|[SourceFullName](http://msdn.microsoft.com/library/linkformat-sourcefullname-property-powerpoint%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/object-model-powerpoint-vba-reference%28Office.15%29.aspx)
