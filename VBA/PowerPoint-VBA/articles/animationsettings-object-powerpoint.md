---
title: AnimationSettings Object (PowerPoint)
keywords: vbapp10.chm565000
f1_keywords:
- vbapp10.chm565000
ms.prod: POWERPOINT
ms.assetid: ebbe4257-236b-35b4-bdf1-e92a1b4b417b
---


# AnimationSettings Object (PowerPoint)

Represents the special effects applied to the animation for the specified shape during a slide show.


## Example

Use the [AnimationSettings](http://msdn.microsoft.com/library/shape-animationsettings-property-powerpoint%28Office.15%29.aspx)property of the  **Shape** object to return the **AnimationSettings** object. The following example adds a slide that contains both a title and a three-item list to the active presentation, and then it sets the list to be animated by first-level paragraphs, to fly in from the left when animated, to dim to the specified color after being animated, and to animate its items in reverse order.


```
Set sObjs = ActivePresentation.Slides.Add(2, ppLayoutText).Shapes

sObjs.Title.TextFrame.TextRange.Text = "Top Three Reasons"

With sObjs.Placeholders(2)

    .TextFrame.TextRange.Text = _

        "Reason 1" &amp; VBNewLine &amp; "Reason 2" &amp; VBNewLine &amp; "Reason 3"

    With .AnimationSettings

        .TextLevelEffect = ppAnimateByFirstLevel

        .EntryEffect = ppEffectFlyFromLeft

        .AfterEffect = ppAfterEffectDim

        .DimColor.RGB = RGB(100, 120, 100)

        .AnimateTextInReverse = True

    End With

End With
```


## Properties



|**Name**|
|:-----|
|[AdvanceMode](http://msdn.microsoft.com/library/animationsettings-advancemode-property-powerpoint%28Office.15%29.aspx)|
|[AdvanceTime](http://msdn.microsoft.com/library/animationsettings-advancetime-property-powerpoint%28Office.15%29.aspx)|
|[AfterEffect](http://msdn.microsoft.com/library/animationsettings-aftereffect-property-powerpoint%28Office.15%29.aspx)|
|[Animate](http://msdn.microsoft.com/library/animationsettings-animate-property-powerpoint%28Office.15%29.aspx)|
|[AnimateBackground](http://msdn.microsoft.com/library/animationsettings-animatebackground-property-powerpoint%28Office.15%29.aspx)|
|[AnimateTextInReverse](http://msdn.microsoft.com/library/animationsettings-animatetextinreverse-property-powerpoint%28Office.15%29.aspx)|
|[AnimationOrder](http://msdn.microsoft.com/library/animationsettings-animationorder-property-powerpoint%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/animationsettings-application-property-powerpoint%28Office.15%29.aspx)|
|[ChartUnitEffect](http://msdn.microsoft.com/library/animationsettings-chartuniteffect-property-powerpoint%28Office.15%29.aspx)|
|[DimColor](http://msdn.microsoft.com/library/animationsettings-dimcolor-property-powerpoint%28Office.15%29.aspx)|
|[EntryEffect](http://msdn.microsoft.com/library/animationsettings-entryeffect-property-powerpoint%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/animationsettings-parent-property-powerpoint%28Office.15%29.aspx)|
|[PlaySettings](http://msdn.microsoft.com/library/animationsettings-playsettings-property-powerpoint%28Office.15%29.aspx)|
|[SoundEffect](http://msdn.microsoft.com/library/animationsettings-soundeffect-property-powerpoint%28Office.15%29.aspx)|
|[TextLevelEffect](http://msdn.microsoft.com/library/animationsettings-textleveleffect-property-powerpoint%28Office.15%29.aspx)|
|[TextUnitEffect](http://msdn.microsoft.com/library/animationsettings-textuniteffect-property-powerpoint%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/object-model-powerpoint-vba-reference%28Office.15%29.aspx)
