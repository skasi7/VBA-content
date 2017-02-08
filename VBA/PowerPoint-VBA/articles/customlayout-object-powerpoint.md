---
title: CustomLayout Object (PowerPoint)
keywords: vbapp10.chm672000
f1_keywords:
- vbapp10.chm672000
ms.prod: POWERPOINT
ms.assetid: 67829704-0314-aed2-5415-6736cefc197e
---


# CustomLayout Object (PowerPoint)

Represents a custom layout associated with a presentation design. The  **CustomLayout** object is a member of the **[CustomLayouts](customlayouts-object-powerpoint.md)** collection.


## Remarks

Use the  **CustomLayout** property of the **[Slide](slide-object-powerpoint.md)** or **[SlideRange](http://msdn.microsoft.com/library/sliderange-object-powerpoint%28Office.15%29.aspx)** objects to access a **CustomLayout** object, for example:


```
ActiveWindow.Selection.SlideRange(1).CustomLayout
```


```
ActivePresentation.Slides(1).CustomLayout
```

Use the  **[Add](http://msdn.microsoft.com/library/customlayouts-add-method-powerpoint%28Office.15%29.aspx)** method of the **CustomLayouts** collection to add a new custom layout to the presentation design's custom layouts. Use the **[Item](http://msdn.microsoft.com/library/customlayouts-add-method-powerpoint%28Office.15%29.aspx)** method to refer to a custom layout. Use the **[Paste](http://msdn.microsoft.com/library/customlayouts-paste-method-powerpoint%28Office.15%29.aspx)** method to paste the slides on the Clipboard into a custom layout and add the custom layout to the **CustomLayouts** collection.


## Methods



|**Name**|
|:-----|
|[Copy](http://msdn.microsoft.com/library/customlayout-copy-method-powerpoint%28Office.15%29.aspx)|
|[Cut](http://msdn.microsoft.com/library/customlayout-cut-method-powerpoint%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/customlayout-delete-method-powerpoint%28Office.15%29.aspx)|
|[Duplicate](http://msdn.microsoft.com/library/customlayout-duplicate-method-powerpoint%28Office.15%29.aspx)|
|[MoveTo](http://msdn.microsoft.com/library/customlayout-moveto-method-powerpoint%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/customlayout-select-method-powerpoint%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/customlayout-application-property-powerpoint%28Office.15%29.aspx)|
|[Background](http://msdn.microsoft.com/library/customlayout-background-property-powerpoint%28Office.15%29.aspx)|
|[CustomerData](http://msdn.microsoft.com/library/customlayout-customerdata-property-powerpoint%28Office.15%29.aspx)|
|[Design](http://msdn.microsoft.com/library/customlayout-design-property-powerpoint%28Office.15%29.aspx)|
|[DisplayMasterShapes](http://msdn.microsoft.com/library/customlayout-displaymastershapes-property-powerpoint%28Office.15%29.aspx)|
|[FollowMasterBackground](http://msdn.microsoft.com/library/customlayout-followmasterbackground-property-powerpoint%28Office.15%29.aspx)|
|[Guides](http://msdn.microsoft.com/library/customlayout-guides-property-powerpoint%28Office.15%29.aspx)|
|[HeadersFooters](http://msdn.microsoft.com/library/customlayout-headersfooters-property-powerpoint%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/customlayout-height-property-powerpoint%28Office.15%29.aspx)|
|[Hyperlinks](http://msdn.microsoft.com/library/customlayout-hyperlinks-property-powerpoint%28Office.15%29.aspx)|
|[Index](http://msdn.microsoft.com/library/customlayout-index-property-powerpoint%28Office.15%29.aspx)|
|[MatchingName](http://msdn.microsoft.com/library/customlayout-matchingname-property-powerpoint%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/customlayout-name-property-powerpoint%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/customlayout-parent-property-powerpoint%28Office.15%29.aspx)|
|[Preserved](http://msdn.microsoft.com/library/customlayout-preserved-property-powerpoint%28Office.15%29.aspx)|
|[Shapes](http://msdn.microsoft.com/library/customlayout-shapes-property-powerpoint%28Office.15%29.aspx)|
|[SlideShowTransition](http://msdn.microsoft.com/library/customlayout-slideshowtransition-property-powerpoint%28Office.15%29.aspx)|
|[ThemeColorScheme](http://msdn.microsoft.com/library/customlayout-themecolorscheme-property-powerpoint%28Office.15%29.aspx)|
|[TimeLine](http://msdn.microsoft.com/library/customlayout-timeline-property-powerpoint%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/customlayout-width-property-powerpoint%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/object-model-powerpoint-vba-reference%28Office.15%29.aspx)
