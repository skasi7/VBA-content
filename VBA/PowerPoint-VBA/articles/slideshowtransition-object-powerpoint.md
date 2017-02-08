---
title: SlideShowTransition Object (PowerPoint)
keywords: vbapp10.chm539000
f1_keywords:
- vbapp10.chm539000
ms.prod: POWERPOINT
ms.assetid: 60707d0d-62a8-0366-c22f-c5c5635fd762
---


# SlideShowTransition Object (PowerPoint)

Contains information about how the specified slide advances during a slide show.


## Example

Use the [SlideShowTransition](http://msdn.microsoft.com/library/slide-slideshowtransition-property-powerpoint%28Office.15%29.aspx)property to return the  **SlideShowTransition** object. The following example specifies a Fast Strips Down-Left transition accompanied by the Bass.wav sound for slide one in the active presentation and specifies that the slide advance automatically five seconds after the previous animation or slide transition.


```
With ActivePresentation.Slides(1).SlideShowTransition

    .Speed = ppTransitionSpeedFast

    .EntryEffect = ppEffectStripsDownLeft

    .SoundEffect.ImportFromFile "c:\sndsys\bass.wav"

    .AdvanceOnTime = True

    .AdvanceTime = 5

End With

ActivePresentation.SlideShowSettings.AdvanceMode = _

    ppSlideShowUseSlideTimings
```


## Properties



|**Name**|
|:-----|
|[AdvanceOnClick](http://msdn.microsoft.com/library/slideshowtransition-advanceonclick-property-powerpoint%28Office.15%29.aspx)|
|[AdvanceOnTime](http://msdn.microsoft.com/library/slideshowtransition-advanceontime-property-powerpoint%28Office.15%29.aspx)|
|[AdvanceTime](http://msdn.microsoft.com/library/slideshowtransition-advancetime-property-powerpoint%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/slideshowtransition-application-property-powerpoint%28Office.15%29.aspx)|
|[Duration](http://msdn.microsoft.com/library/slideshowtransition-duration-property-powerpoint%28Office.15%29.aspx)|
|[EntryEffect](http://msdn.microsoft.com/library/slideshowtransition-entryeffect-property-powerpoint%28Office.15%29.aspx)|
|[Hidden](http://msdn.microsoft.com/library/slideshowtransition-hidden-property-powerpoint%28Office.15%29.aspx)|
|[LoopSoundUntilNext](http://msdn.microsoft.com/library/slideshowtransition-loopsounduntilnext-property-powerpoint%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/slideshowtransition-parent-property-powerpoint%28Office.15%29.aspx)|
|[SoundEffect](http://msdn.microsoft.com/library/slideshowtransition-soundeffect-property-powerpoint%28Office.15%29.aspx)|
|[Speed](http://msdn.microsoft.com/library/slideshowtransition-speed-property-powerpoint%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/object-model-powerpoint-vba-reference%28Office.15%29.aspx)
