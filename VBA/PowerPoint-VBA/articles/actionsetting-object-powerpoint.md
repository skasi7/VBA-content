---
title: ActionSetting Object (PowerPoint)
keywords: vbapp10.chm567000
f1_keywords:
- vbapp10.chm567000
ms.prod: POWERPOINT
ms.assetid: 21381ff0-b9ff-59d8-77e9-345905fb8617
---


# ActionSetting Object (PowerPoint)

Contains information about how the specified shape or text range reacts to mouse actions during a slide show. 


## Remarks

The  **ActionSetting** object is a member of the **[ActionSettings](http://msdn.microsoft.com/library/actionsettings-object-powerpoint%28Office.15%29.aspx)** collection. The **ActionSettings** collection contains one **ActionSetting** object that represents how the specified object reacts when the user clicks it during a slide show and one **ActionSetting** object that represents how the specified object reacts when the user moves the mouse pointer over it during a slide show.

If you've set properties of the  **ActionSetting** object that don't seem to be taking effect, make sure that you've set the[Action](http://msdn.microsoft.com/library/actionsetting-action-property-powerpoint%28Office.15%29.aspx) property to the appropriate value.


## Example

Use  **ActionSettings** (index), where index is the either **ppMouseClick** or **ppMouseOver**, to return a single **ActionSetting** object. The following example sets the mouse-click action for the text in the third shape on slide one in the active presentation to an Internet link.


```
With ActivePresentation.Slides(1).Shapes(3) _ 
        .TextFrame.TextRange.ActionSettings(ppMouseClick) 
    .Action = ppActionHyperlink 
    .Hyperlink.Address = "http://www.microsoft.com" 
End With
```


## Properties



|**Name**|
|:-----|
|[Action](http://msdn.microsoft.com/library/actionsetting-action-property-powerpoint%28Office.15%29.aspx)|
|[ActionVerb](http://msdn.microsoft.com/library/actionsetting-actionverb-property-powerpoint%28Office.15%29.aspx)|
|[AnimateAction](http://msdn.microsoft.com/library/actionsetting-animateaction-property-powerpoint%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/actionsetting-application-property-powerpoint%28Office.15%29.aspx)|
|[Hyperlink](http://msdn.microsoft.com/library/actionsetting-hyperlink-property-powerpoint%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/actionsetting-parent-property-powerpoint%28Office.15%29.aspx)|
|[Run](http://msdn.microsoft.com/library/actionsetting-run-property-powerpoint%28Office.15%29.aspx)|
|[ShowAndReturn](http://msdn.microsoft.com/library/actionsetting-showandreturn-property-powerpoint%28Office.15%29.aspx)|
|[SlideShowName](http://msdn.microsoft.com/library/actionsetting-slideshowname-property-powerpoint%28Office.15%29.aspx)|
|[SoundEffect](http://msdn.microsoft.com/library/actionsetting-soundeffect-property-powerpoint%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/object-model-powerpoint-vba-reference%28Office.15%29.aspx)
