---
title: ParagraphFormat Object (PowerPoint)
keywords: vbapp10.chm576000
f1_keywords:
- vbapp10.chm576000
ms.prod: POWERPOINT
ms.assetid: 15d495cf-16e2-5cfb-e99c-a551876e3a8a
---


# ParagraphFormat Object (PowerPoint)

Represents the paragraph formatting of a text range.


## Example

Use the [ParagraphFormat](http://msdn.microsoft.com/library/textrange-paragraphformat-property-powerpoint%28Office.15%29.aspx)property to return the  **ParagraphFormat** object. The following example left aligns the paragraphs in shape two on slide one in the active presentation.


```
ActivePresentation.Slides(1).Shapes(2).TextFrame.TextRange _

    .ParagraphFormat.Alignment = ppAlignLeft
```


## Properties



|**Name**|
|:-----|
|[Alignment](http://msdn.microsoft.com/library/paragraphformat-alignment-property-powerpoint%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/paragraphformat-application-property-powerpoint%28Office.15%29.aspx)|
|[BaseLineAlignment](http://msdn.microsoft.com/library/paragraphformat-baselinealignment-property-powerpoint%28Office.15%29.aspx)|
|[Bullet](http://msdn.microsoft.com/library/paragraphformat-bullet-property-powerpoint%28Office.15%29.aspx)|
|[FarEastLineBreakControl](http://msdn.microsoft.com/library/paragraphformat-fareastlinebreakcontrol-property-powerpoint%28Office.15%29.aspx)|
|[HangingPunctuation](http://msdn.microsoft.com/library/paragraphformat-hangingpunctuation-property-powerpoint%28Office.15%29.aspx)|
|[LineRuleAfter](http://msdn.microsoft.com/library/paragraphformat-lineruleafter-property-powerpoint%28Office.15%29.aspx)|
|[LineRuleBefore](http://msdn.microsoft.com/library/paragraphformat-linerulebefore-property-powerpoint%28Office.15%29.aspx)|
|[LineRuleWithin](http://msdn.microsoft.com/library/paragraphformat-linerulewithin-property-powerpoint%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/paragraphformat-parent-property-powerpoint%28Office.15%29.aspx)|
|[SpaceAfter](http://msdn.microsoft.com/library/paragraphformat-spaceafter-property-powerpoint%28Office.15%29.aspx)|
|[SpaceBefore](http://msdn.microsoft.com/library/paragraphformat-spacebefore-property-powerpoint%28Office.15%29.aspx)|
|[SpaceWithin](http://msdn.microsoft.com/library/paragraphformat-spacewithin-property-powerpoint%28Office.15%29.aspx)|
|[TextDirection](http://msdn.microsoft.com/library/paragraphformat-textdirection-property-powerpoint%28Office.15%29.aspx)|
|[WordWrap](http://msdn.microsoft.com/library/paragraphformat-wordwrap-property-powerpoint%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/object-model-powerpoint-vba-reference%28Office.15%29.aspx)
