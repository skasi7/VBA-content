---
title: SlideShowWindow Object (PowerPoint)
keywords: vbapp10.chm507000
f1_keywords:
- vbapp10.chm507000
ms.prod: POWERPOINT
ms.assetid: 22468489-d4a2-ffea-7479-53ecb8d5da29
---


# SlideShowWindow Object (PowerPoint)

Represents a window in which a slide show runs.


## Example

Use  **SlideShowWindows** (index), where index is the slide show window index number, to return a single **SlideShowWindow** object. The following example activates slide show window two.


```
SlideShowWindows(2).Activate
```

Use the [Run](http://msdn.microsoft.com/library/slideshowsettings-run-method-powerpoint%28Office.15%29.aspx)method to create a new slide show window and return a reference to this slide show window. The following example runs a slide show of the active presentation and reduces the height of the slide show window just enough so that you can see the taskbar (for monitors with a screen resolution of 800 by 600).




```
With ActivePresentation.SlideShowSettings

    .ShowType = ppShowTypeSpeaker

    With .Run

        .Height = 300

        .Width = 400

    End With

End With
```

Use the [View](http://msdn.microsoft.com/library/slideshowwindow-view-property-powerpoint%28Office.15%29.aspx)property to return the view in the specified slide show window. The following example sets the view in slide show window one to display slide three in the presentation.




```
SlideShowWindows(1).View.GotoSlide 3
```

Use the [Presentation](http://msdn.microsoft.com/library/slideshowwindow-presentation-property-powerpoint%28Office.15%29.aspx)property to return the presentation that's currently running in the specified slide show window. The following example displays the name of the presentation that's currently running in slide show window one.




```
MsgBox SlideShowWindows(1).Presentation.Name
```


## Methods



|**Name**|
|:-----|
|[DrawLine](http://msdn.microsoft.com/library/slideshowview-drawline-method-powerpoint%28Office.15%29.aspx)|
|[EndNamedShow](http://msdn.microsoft.com/library/slideshowview-endnamedshow-method-powerpoint%28Office.15%29.aspx)|
|[EraseDrawing](http://msdn.microsoft.com/library/slideshowview-erasedrawing-method-powerpoint%28Office.15%29.aspx)|
|[Exit](http://msdn.microsoft.com/library/slideshowview-exit-method-powerpoint%28Office.15%29.aspx)|
|[First](http://msdn.microsoft.com/library/slideshowview-first-method-powerpoint%28Office.15%29.aspx)|
|[FirstAnimationIsAutomatic](http://msdn.microsoft.com/library/slideshowview-firstanimationisautomatic-method-powerpoint%28Office.15%29.aspx)|
|[GetClickCount](http://msdn.microsoft.com/library/slideshowview-getclickcount-method-powerpoint%28Office.15%29.aspx)|
|[GetClickIndex](http://msdn.microsoft.com/library/slideshowview-getclickindex-method-powerpoint%28Office.15%29.aspx)|
|[GotoClick](http://msdn.microsoft.com/library/slideshowview-gotoclick-method-powerpoint%28Office.15%29.aspx)|
|[GotoNamedShow](http://msdn.microsoft.com/library/slideshowview-gotonamedshow-method-powerpoint%28Office.15%29.aspx)|
|[GotoSlide](http://msdn.microsoft.com/library/slideshowview-gotoslide-method-powerpoint%28Office.15%29.aspx)|
|[Last](http://msdn.microsoft.com/library/slideshowview-last-method-powerpoint%28Office.15%29.aspx)|
|[Next](http://msdn.microsoft.com/library/slideshowview-next-method-powerpoint%28Office.15%29.aspx)|
|[Player](http://msdn.microsoft.com/library/slideshowview-player-method-powerpoint%28Office.15%29.aspx)|
|[Previous](http://msdn.microsoft.com/library/slideshowview-previous-method-powerpoint%28Office.15%29.aspx)|
|[ResetSlideTime](http://msdn.microsoft.com/library/slideshowview-resetslidetime-method-powerpoint%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AcceleratorsEnabled](http://msdn.microsoft.com/library/slideshowview-acceleratorsenabled-property-powerpoint%28Office.15%29.aspx)|
|[AdvanceMode](http://msdn.microsoft.com/library/slideshowview-advancemode-property-powerpoint%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/slideshowview-application-property-powerpoint%28Office.15%29.aspx)|
|[CurrentShowPosition](http://msdn.microsoft.com/library/slideshowview-currentshowposition-property-powerpoint%28Office.15%29.aspx)|
|[IsNamedShow](http://msdn.microsoft.com/library/slideshowview-isnamedshow-property-powerpoint%28Office.15%29.aspx)|
|[LaserPointerEnabled](http://msdn.microsoft.com/library/slideshowview-laserpointerenabled-property-powerpoint%28Office.15%29.aspx)|
|[LastSlideViewed](http://msdn.microsoft.com/library/slideshowview-lastslideviewed-property-powerpoint%28Office.15%29.aspx)|
|[MediaControlsHeight](http://msdn.microsoft.com/library/slideshowview-mediacontrolsheight-property-powerpoint%28Office.15%29.aspx)|
|[MediaControlsLeft](http://msdn.microsoft.com/library/slideshowview-mediacontrolsleft-property-powerpoint%28Office.15%29.aspx)|
|[MediaControlsTop](http://msdn.microsoft.com/library/slideshowview-mediacontrolstop-property-powerpoint%28Office.15%29.aspx)|
|[MediaControlsVisible](http://msdn.microsoft.com/library/slideshowview-mediacontrolsvisible-property-powerpoint%28Office.15%29.aspx)|
|[MediaControlsWidth](http://msdn.microsoft.com/library/slideshowview-mediacontrolswidth-property-powerpoint%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/slideshowview-parent-property-powerpoint%28Office.15%29.aspx)|
|[PointerColor](http://msdn.microsoft.com/library/slideshowview-pointercolor-property-powerpoint%28Office.15%29.aspx)|
|[PointerType](http://msdn.microsoft.com/library/slideshowview-pointertype-property-powerpoint%28Office.15%29.aspx)|
|[PresentationElapsedTime](http://msdn.microsoft.com/library/slideshowview-presentationelapsedtime-property-powerpoint%28Office.15%29.aspx)|
|[Slide](http://msdn.microsoft.com/library/slideshowview-slide-property-powerpoint%28Office.15%29.aspx)|
|[SlideElapsedTime](http://msdn.microsoft.com/library/slideshowview-slideelapsedtime-property-powerpoint%28Office.15%29.aspx)|
|[SlideShowName](http://msdn.microsoft.com/library/slideshowview-slideshowname-property-powerpoint%28Office.15%29.aspx)|
|[State](http://msdn.microsoft.com/library/slideshowview-state-property-powerpoint%28Office.15%29.aspx)|
|[Zoom](http://msdn.microsoft.com/library/slideshowview-zoom-property-powerpoint%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/object-model-powerpoint-vba-reference%28Office.15%29.aspx)
