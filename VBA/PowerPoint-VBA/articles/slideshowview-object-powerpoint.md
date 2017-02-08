---
title: SlideShowView Object (PowerPoint)
keywords: vbapp10.chm513000
f1_keywords:
- vbapp10.chm513000
ms.prod: POWERPOINT
ms.assetid: 403b30ef-b12f-3a3c-e8d8-19189fd762fe
---


# SlideShowView Object (PowerPoint)

Represents the view in a slide show window.


## Example

Use the [View](http://msdn.microsoft.com/library/slideshowwindow-view-property-powerpoint%28Office.15%29.aspx)property of the  **SlideShowWindow** object to return the **SlideShowView** object. The following example sets slide show window one to display the first slide in the presentation.


```
SlideShowWindows(1).View.First
```

Use the [Run](http://msdn.microsoft.com/library/slideshowsettings-run-method-powerpoint%28Office.15%29.aspx)method of the  **SlideShowSettings** object to create a **SlideShowWindow** object, and then use the **View** property to return the **SlideShowView** object the window contains. The following example runs a slide show of the active presentation, changes the pointer to a pen, and sets the pen color for the slide show to red.




```
With ActivePresentation.SlideShowSettings.Run.View

    .PointerColor.RGB = RGB(255, 0, 0)

    .PointerType = ppSlideShowPointerPen

End With
```


## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/object-model-powerpoint-vba-reference%28Office.15%29.aspx)
[SlideShowView Object Members](http://msdn.microsoft.com/library/slideshowview-members-powerpoint%28Office.15%29.aspx)
