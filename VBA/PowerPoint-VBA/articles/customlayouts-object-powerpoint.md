---
title: CustomLayouts Object (PowerPoint)
keywords: vbapp10.chm671000
f1_keywords:
- vbapp10.chm671000
ms.prod: POWERPOINT
ms.assetid: 9ce682fb-545c-55cb-e9ac-3475f7556af1
---


# CustomLayouts Object (PowerPoint)

Represents a set of custom layouts associated with a presentation design.


## Remarks

Use the  **[CustomLayouts](http://msdn.microsoft.com/library/master-customlayouts-property-powerpoint%28Office.15%29.aspx)** property of the slide **[Master](master-object-powerpoint.md)** object to return a **CustomLayouts** collection. Use **CustomLayouts** ( _index_ ), where index is the color scheme index number, to return a single **[CustomLayout](customlayout-object-powerpoint.md)** object.

Use the  **[Add](http://msdn.microsoft.com/library/customlayouts-add-method-powerpoint%28Office.15%29.aspx)** method to create a new custom layout and add it to the **CustomLayouts** collection. Use the **[Paste](http://msdn.microsoft.com/library/customlayouts-paste-method-powerpoint%28Office.15%29.aspx)** method to past slides from the Clipboard as a **CustomLayout** object into the **CustomLayouts** collection.

Use the  **CustomLayout** property of a **[Slide](slide-object-powerpoint.md)** or **[SlideRange](http://msdn.microsoft.com/library/sliderange-object-powerpoint%28Office.15%29.aspx)** object to return a custom layout for a slide or set of slides.


## Example

The following example adds a custom layout to the slide master of the active presentation.


```
Sub AddCustomLayout()

    With ActivePresentation.SlideMaster

        .CustomLayouts.Add (1)

        .CustomLayouts(1).Name = "MyLayout"

    End With

End Sub
```

The following example displays the name of the custom layout for the first slide of the active presentation.




```
MsgBox ActivePresentation.Slides(1).CustomLayout.Name
```


## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/customlayouts-add-method-powerpoint%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/customlayouts-item-method-powerpoint%28Office.15%29.aspx)|
|[Paste](http://msdn.microsoft.com/library/customlayouts-paste-method-powerpoint%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/customlayouts-application-property-powerpoint%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/customlayouts-count-property-powerpoint%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/customlayouts-parent-property-powerpoint%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/object-model-powerpoint-vba-reference%28Office.15%29.aspx)
