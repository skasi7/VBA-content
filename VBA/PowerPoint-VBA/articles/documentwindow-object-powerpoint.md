---
title: DocumentWindow Object (PowerPoint)
keywords: vbapp10.chm511000
f1_keywords:
- vbapp10.chm511000
ms.prod: POWERPOINT
ms.assetid: 567c5e66-8d68-a868-4072-b5358cf69546
---


# DocumentWindow Object (PowerPoint)

Represents a document window. The  **DocumentWindow** object is a member of the **[DocumentWindows](http://msdn.microsoft.com/library/documentwindows-object-powerpoint%28Office.15%29.aspx)** collection. The **DocumentWindows** collection contains all the open document windows.


## Remarks

Use the  **[Presentation](http://msdn.microsoft.com/library/application-presentations-property-powerpoint%28Office.15%29.aspx)** property to return the presentation that's currently running in the specified document window.

Use the  **[Selection](http://msdn.microsoft.com/library/documentwindow-selection-property-powerpoint%28Office.15%29.aspx)** property to return the selection.

Use the  **[SplitHorizontal](http://msdn.microsoft.com/library/documentwindow-splithorizontal-property-powerpoint%28Office.15%29.aspx)** property to return the percentage of the screen width that the outline pane occupies in normal view.

Use the  **[SplitVertical](http://msdn.microsoft.com/library/documentwindow-splitvertical-property-powerpoint%28Office.15%29.aspx)** property to return the percentage of the screen height that the slide pane occupies in normal view.

Use the  **[View](http://msdn.microsoft.com/library/documentwindow-view-property-powerpoint%28Office.15%29.aspx)** property to return the view in the specified document window.


## Example

Use  **Windows** (index), where index is the document window index number, to return a single **DocumentWindow** object. The following example activates document window two.


```
Windows(2).Activate
```

The first member of the  **DocumentWindows** collection, `Windows(1)`, always returns the active document window. Alternatively, you can use the  **[ActiveWindow](http://msdn.microsoft.com/library/application-activewindow-property-powerpoint%28Office.15%29.aspx)** property to return the active document window. The following example maximizes the active window.




```
ActiveWindow.WindowState = ppWindowMaximized
```

Use  **Panes** (index), where index is the pane index number, to manipulate panes within normal, slide, outline, or notes page views of the document window. The following example activates pane three, which is the notes pane.




```
ActiveWindow.Panes(3).Activate
```

Use the  **[ActivePane](http://msdn.microsoft.com/library/documentwindow-activepane-property-powerpoint%28Office.15%29.aspx)** property to return the active pane within the document window. The following example checks to see if the active pane is the outline pane. If not, it activates the outline pane.




```
mypane = ActiveWindow.ActivePane.ViewType

    If mypane <> 1 Then

        ActiveWindow.Panes(1).Activate

    End If
```


## Methods



|**Name**|
|:-----|
|**[Activate](http://msdn.microsoft.com/library/documentwindow-activate-method-powerpoint%28Office.15%29.aspx)**|
|**[Close](http://msdn.microsoft.com/library/documentwindow-close-method-powerpoint%28Office.15%29.aspx)**|
|**[ExpandSection](http://msdn.microsoft.com/library/documentwindow-expandsection-method-powerpoint%28Office.15%29.aspx)**|
|**[FitToPage](http://msdn.microsoft.com/library/documentwindow-fittopage-method-powerpoint%28Office.15%29.aspx)**|
|**[IsSectionExpanded](http://msdn.microsoft.com/library/documentwindow-issectionexpanded-method-powerpoint%28Office.15%29.aspx)**|
|**[LargeScroll](http://msdn.microsoft.com/library/documentwindow-largescroll-method-powerpoint%28Office.15%29.aspx)**|
|**[NewWindow](http://msdn.microsoft.com/library/documentwindow-newwindow-method-powerpoint%28Office.15%29.aspx)**|
|**[PointsToScreenPixelsX](http://msdn.microsoft.com/library/documentwindow-pointstoscreenpixelsx-method-powerpoint%28Office.15%29.aspx)**|
|**[PointsToScreenPixelsY](http://msdn.microsoft.com/library/documentwindow-pointstoscreenpixelsy-method-powerpoint%28Office.15%29.aspx)**|
|**[RangeFromPoint](http://msdn.microsoft.com/library/documentwindow-rangefrompoint-method-powerpoint%28Office.15%29.aspx)**|
|**[ScrollIntoView](http://msdn.microsoft.com/library/documentwindow-scrollintoview-method-powerpoint%28Office.15%29.aspx)**|
|**[SmallScroll](http://msdn.microsoft.com/library/documentwindow-smallscroll-method-powerpoint%28Office.15%29.aspx)**|

## Properties



|**Name**|
|:-----|
|**[Active](http://msdn.microsoft.com/library/documentwindow-active-property-powerpoint%28Office.15%29.aspx)**|
|**[ActivePane](http://msdn.microsoft.com/library/documentwindow-activepane-property-powerpoint%28Office.15%29.aspx)**|
|**[Application](http://msdn.microsoft.com/library/documentwindow-application-property-powerpoint%28Office.15%29.aspx)**|
|**[BlackAndWhite](http://msdn.microsoft.com/library/documentwindow-blackandwhite-property-powerpoint%28Office.15%29.aspx)**|
|**[Caption](http://msdn.microsoft.com/library/documentwindow-caption-property-powerpoint%28Office.15%29.aspx)**|
|**[Height](http://msdn.microsoft.com/library/documentwindow-height-property-powerpoint%28Office.15%29.aspx)**|
|**[Left](http://msdn.microsoft.com/library/documentwindow-left-property-powerpoint%28Office.15%29.aspx)**|
|**[Panes](http://msdn.microsoft.com/library/documentwindow-panes-property-powerpoint%28Office.15%29.aspx)**|
|**[Parent](http://msdn.microsoft.com/library/documentwindow-parent-property-powerpoint%28Office.15%29.aspx)**|
|**[Presentation](http://msdn.microsoft.com/library/documentwindow-presentation-property-powerpoint%28Office.15%29.aspx)**|
|**[Selection](http://msdn.microsoft.com/library/documentwindow-selection-property-powerpoint%28Office.15%29.aspx)**|
|**[SplitHorizontal](http://msdn.microsoft.com/library/documentwindow-splithorizontal-property-powerpoint%28Office.15%29.aspx)**|
|**[SplitVertical](http://msdn.microsoft.com/library/documentwindow-splitvertical-property-powerpoint%28Office.15%29.aspx)**|
|**[Top](http://msdn.microsoft.com/library/documentwindow-top-property-powerpoint%28Office.15%29.aspx)**|
|**[View](http://msdn.microsoft.com/library/documentwindow-view-property-powerpoint%28Office.15%29.aspx)**|
|**[ViewType](http://msdn.microsoft.com/library/documentwindow-viewtype-property-powerpoint%28Office.15%29.aspx)**|
|**[Width](http://msdn.microsoft.com/library/documentwindow-width-property-powerpoint%28Office.15%29.aspx)**|
|**[WindowState](http://msdn.microsoft.com/library/documentwindow-windowstate-property-powerpoint%28Office.15%29.aspx)**|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/object-model-powerpoint-vba-reference%28Office.15%29.aspx)
