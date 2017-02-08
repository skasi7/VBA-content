---
title: TimeLine Object (PowerPoint)
keywords: vbapp10.chm649000
f1_keywords:
- vbapp10.chm649000
ms.prod: POWERPOINT
ms.assetid: 0b5a8863-8329-48d0-cb0b-3b34e87acb76
---


# TimeLine Object (PowerPoint)

Stores animation information for a  **Master**, **Slide**, or **SlideRange** object.


## Example

Use the [TimeLine](http://msdn.microsoft.com/library/master-timeline-property-powerpoint%28Office.15%29.aspx)property of the  **[Master](master-object-powerpoint.md)**, **[Slide](slide-object-powerpoint.md)**, or **[SlideRange](http://msdn.microsoft.com/library/sliderange-object-powerpoint%28Office.15%29.aspx)** object to return a **TimeLine** object.

The  **TimeLine** object's **[MainSequence](http://msdn.microsoft.com/library/timeline-mainsequence-property-powerpoint%28Office.15%29.aspx)** property gains access to the main animation sequence, while the **[InteractiveSequences](http://msdn.microsoft.com/library/timeline-interactivesequences-property-powerpoint%28Office.15%29.aspx)** property gains access to the collection of interactive animation sequences of a slide or slide range.

To reference a timeline object, use syntax similar to these code examples:




```
ActivePresentation.Slides(1).TimeLine.MainSequence

ActivePresentation.SlideMaster.TimeLine.InteractiveSequences

ActiveWindow.Selection.SlideRange.TimeLine.InteractiveSequences
```


## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/timeline-application-property-powerpoint%28Office.15%29.aspx)|
|[InteractiveSequences](http://msdn.microsoft.com/library/timeline-interactivesequences-property-powerpoint%28Office.15%29.aspx)|
|[MainSequence](http://msdn.microsoft.com/library/timeline-mainsequence-property-powerpoint%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/timeline-parent-property-powerpoint%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/object-model-powerpoint-vba-reference%28Office.15%29.aspx)
