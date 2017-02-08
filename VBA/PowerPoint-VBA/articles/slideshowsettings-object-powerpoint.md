---
title: SlideShowSettings Object (PowerPoint)
keywords: vbapp10.chm514000
f1_keywords:
- vbapp10.chm514000
ms.prod: POWERPOINT
ms.assetid: d58c7c3b-a1cc-d819-b386-fd3fb7f967a2
---


# SlideShowSettings Object (PowerPoint)

Represents the slide show setup for a presentation.


## Example

Use the [SlideShowSettings](http://msdn.microsoft.com/library/presentation-slideshowsettings-property-powerpoint%28Office.15%29.aspx)property to return the  **SlideShowSettings** object. The first section in the following example sets all the slides in the active presentation to advance automatically after five seconds. The second section sets the slide show to start on slide two, end on slide four, advance slides by using the timings set in the first section, and run in a continuous loop until the user presses ESC. Finally, the example runs the slide show.


```
For Each s In ActivePresentation.Slides

    With s.SlideShowTransition

        .AdvanceOnTime = True

        .AdvanceTime = 5

    End With

Next



With ActivePresentation.SlideShowSettings

    .RangeType = ppShowSlideRange

    .StartingSlide = 2

    .EndingSlide = 4

    .AdvanceMode = ppSlideShowUseSlideTimings

    .LoopUntilStopped = True

    .Run

End With
```


## Methods



|**Name**|
|:-----|
|[Run](http://msdn.microsoft.com/library/slideshowsettings-run-method-powerpoint%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AdvanceMode](http://msdn.microsoft.com/library/slideshowsettings-advancemode-property-powerpoint%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/slideshowsettings-application-property-powerpoint%28Office.15%29.aspx)|
|[EndingSlide](http://msdn.microsoft.com/library/slideshowsettings-endingslide-property-powerpoint%28Office.15%29.aspx)|
|[LoopUntilStopped](http://msdn.microsoft.com/library/slideshowsettings-loopuntilstopped-property-powerpoint%28Office.15%29.aspx)|
|[NamedSlideShows](http://msdn.microsoft.com/library/slideshowsettings-namedslideshows-property-powerpoint%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/slideshowsettings-parent-property-powerpoint%28Office.15%29.aspx)|
|[PointerColor](http://msdn.microsoft.com/library/slideshowsettings-pointercolor-property-powerpoint%28Office.15%29.aspx)|
|[RangeType](http://msdn.microsoft.com/library/slideshowsettings-rangetype-property-powerpoint%28Office.15%29.aspx)|
|[ShowMediaControls](http://msdn.microsoft.com/library/slideshowsettings-showmediacontrols-property-powerpoint%28Office.15%29.aspx)|
|[ShowPresenterView](http://msdn.microsoft.com/library/slideshowsettings-showpresenterview-property-powerpoint%28Office.15%29.aspx)|
|[ShowScrollbar](http://msdn.microsoft.com/library/slideshowsettings-showscrollbar-property-powerpoint%28Office.15%29.aspx)|
|[ShowType](http://msdn.microsoft.com/library/slideshowsettings-showtype-property-powerpoint%28Office.15%29.aspx)|
|[ShowWithAnimation](http://msdn.microsoft.com/library/slideshowsettings-showwithanimation-property-powerpoint%28Office.15%29.aspx)|
|[ShowWithNarration](http://msdn.microsoft.com/library/slideshowsettings-showwithnarration-property-powerpoint%28Office.15%29.aspx)|
|[SlideShowName](http://msdn.microsoft.com/library/slideshowsettings-slideshowname-property-powerpoint%28Office.15%29.aspx)|
|[StartingSlide](http://msdn.microsoft.com/library/slideshowsettings-startingslide-property-powerpoint%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/object-model-powerpoint-vba-reference%28Office.15%29.aspx)
