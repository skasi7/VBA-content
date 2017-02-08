---
title: Effect Object (PowerPoint)
keywords: vbapp10.chm652000
f1_keywords:
- vbapp10.chm652000
ms.prod: POWERPOINT
ms.assetid: 359ac3da-86cd-8003-d691-349d20fd1777
---


# Effect Object (PowerPoint)

Represents timing information about a slide animation.


## Example

Use the [AddEffect](http://msdn.microsoft.com/library/sequence-addeffect-method-powerpoint%28Office.15%29.aspx)method to add an effect. This example adds a shape to the first slide in the active presentation and adds an effect and a behavior to the shape.


```
Sub NewShapeAndEffect()

    Dim shpStar As Shape

    Dim sldOne As Slide

    Dim effNew As Effect



    Set sldOne = ActivePresentation.Slides(1)

    Set shpStar = sldOne.Shapes.AddShape(Type:=msoShape5pointStar, _

        Left:=150, Top:=72, Width:=400, Height:=400)

    Set effNew = sldOne.TimeLine.MainSequence.AddEffect(Shape:=shpStar, _

        EffectId:=msoAnimEffectStretchy, Trigger:=msoAnimTriggerAfterPrevious)

    With effNew

        With .Behaviors.Add(msoAnimTypeScale).ScaleEffect

            .FromX = 75

            .FromY = 75

            .ToX = 0

            .ToY = 0

        End With

        .Timing.AutoReverse = msoTrue

    End With

End Sub
```

To refer to an existing  **Effect** object, use **[MainSequence](http://msdn.microsoft.com/library/timeline-mainsequence-property-powerpoint%28Office.15%29.aspx)** (index), where index is the number of the **Effect** object in the **[Sequence](http://msdn.microsoft.com/library/sequence-object-powerpoint%28Office.15%29.aspx)** collection. This example changes the effect for the first sequence and specifies the behavior for that effect.




```
Sub ChangeEffect()

    With ActivePresentation.Slides(1).TimeLine _

        .MainSequence(1)

        .EffectType = msoAnimEffectSpin

        With .Behaviors(1).RotationEffect

            .From = 100

            .To = 360

            .By = 5

        End With

    End With

End Sub
```


## Methods



|**Name**|
|:-----|
|**[Delete](http://msdn.microsoft.com/library/effect-delete-method-powerpoint%28Office.15%29.aspx)**|
|**[MoveAfter](http://msdn.microsoft.com/library/effect-moveafter-method-powerpoint%28Office.15%29.aspx)**|
|**[MoveBefore](http://msdn.microsoft.com/library/effect-movebefore-method-powerpoint%28Office.15%29.aspx)**|
|**[MoveTo](http://msdn.microsoft.com/library/effect-moveto-method-powerpoint%28Office.15%29.aspx)**|

## Properties



|**Name**|
|:-----|
|**[Application](http://msdn.microsoft.com/library/effect-application-property-powerpoint%28Office.15%29.aspx)**|
|**[Behaviors](http://msdn.microsoft.com/library/effect-behaviors-property-powerpoint%28Office.15%29.aspx)**|
|**[DisplayName](http://msdn.microsoft.com/library/effect-displayname-property-powerpoint%28Office.15%29.aspx)**|
|**[EffectInformation](http://msdn.microsoft.com/library/effect-effectinformation-property-powerpoint%28Office.15%29.aspx)**|
|**[EffectParameters](http://msdn.microsoft.com/library/effect-effectparameters-property-powerpoint%28Office.15%29.aspx)**|
|**[EffectType](http://msdn.microsoft.com/library/effect-effecttype-property-powerpoint%28Office.15%29.aspx)**|
|**[Exit](http://msdn.microsoft.com/library/effect-exit-property-powerpoint%28Office.15%29.aspx)**|
|**[Index](http://msdn.microsoft.com/library/effect-index-property-powerpoint%28Office.15%29.aspx)**|
|**[Paragraph](http://msdn.microsoft.com/library/effect-paragraph-property-powerpoint%28Office.15%29.aspx)**|
|**[Parent](http://msdn.microsoft.com/library/effect-parent-property-powerpoint%28Office.15%29.aspx)**|
|**[Shape](http://msdn.microsoft.com/library/effect-shape-property-powerpoint%28Office.15%29.aspx)**|
|**[TextRangeLength](http://msdn.microsoft.com/library/effect-textrangelength-property-powerpoint%28Office.15%29.aspx)**|
|**[TextRangeStart](http://msdn.microsoft.com/library/effect-textrangestart-property-powerpoint%28Office.15%29.aspx)**|
|**[Timing](http://msdn.microsoft.com/library/effect-timing-property-powerpoint%28Office.15%29.aspx)**|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/object-model-powerpoint-vba-reference%28Office.15%29.aspx)
