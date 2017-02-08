---
title: Master Object (PowerPoint)
keywords: vbapp10.chm638000
f1_keywords:
- vbapp10.chm638000
ms.prod: POWERPOINT
ms.assetid: 22e8805e-6469-1a34-7f7b-f1ea5c6c49ff
---


# Master Object (PowerPoint)

Represents a slide master, title master, handout master, notes master, or design master.


## Example

To return a  **Master** object, use the[Master](http://msdn.microsoft.com/library/slide-master-property-powerpoint%28Office.15%29.aspx)property of the  **[Slide](slide-object-powerpoint.md)** object or **[SlideRange](http://msdn.microsoft.com/library/sliderange-object-powerpoint%28Office.15%29.aspx)** collection, or use the[HandoutMaster](http://msdn.microsoft.com/library/presentation-handoutmaster-property-powerpoint%28Office.15%29.aspx), [NotesMaster](http://msdn.microsoft.com/library/presentation-notesmaster-property-powerpoint%28Office.15%29.aspx), [SlideMaster](http://msdn.microsoft.com/library/design-slidemaster-property-powerpoint%28Office.15%29.aspx), or [TitleMaster](http://msdn.microsoft.com/library/presentation-titlemaster-property-powerpoint%28Office.15%29.aspx)property of the  **[Presentation](presentation-object-powerpoint.md)** object. Note that some of these properties are also available from the **[Design](http://msdn.microsoft.com/library/design-object-powerpoint%28Office.15%29.aspx)** object as well. The following example sets the background fill for the slide master for the active presentation.


```
ActivePresentation.SlideMaster.Background.Fill _

    .PresetGradient msoGradientHorizontal, 1, msoGradientBrass
```

To add a title master or design to a presentation and return a  **Master** object that represents the new title master or design, use the[AddTitleMaster](http://msdn.microsoft.com/library/presentation-addtitlemaster-method-powerpoint%28Office.15%29.aspx)method. The following example adds a title master to the active presentation and places the title placeholder 10 points from the top of the master.




```
ActivePresentation.AddTitleMaster.Shapes.Title.Top = 10
```


## Methods



|**Name**|
|:-----|
|[ApplyTheme](http://msdn.microsoft.com/library/master-applytheme-method-powerpoint%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/master-delete-method-powerpoint%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/master-application-property-powerpoint%28Office.15%29.aspx)|
|[Background](http://msdn.microsoft.com/library/master-background-property-powerpoint%28Office.15%29.aspx)|
|[BackgroundStyle](http://msdn.microsoft.com/library/master-backgroundstyle-property-powerpoint%28Office.15%29.aspx)|
|[ColorScheme](http://msdn.microsoft.com/library/master-colorscheme-property-powerpoint%28Office.15%29.aspx)|
|[CustomerData](http://msdn.microsoft.com/library/master-customerdata-property-powerpoint%28Office.15%29.aspx)|
|[CustomLayouts](http://msdn.microsoft.com/library/master-customlayouts-property-powerpoint%28Office.15%29.aspx)|
|[Design](http://msdn.microsoft.com/library/master-design-property-powerpoint%28Office.15%29.aspx)|
|[Guides](http://msdn.microsoft.com/library/master-guides-property-powerpoint%28Office.15%29.aspx)|
|[HeadersFooters](http://msdn.microsoft.com/library/master-headersfooters-property-powerpoint%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/master-height-property-powerpoint%28Office.15%29.aspx)|
|[Hyperlinks](http://msdn.microsoft.com/library/master-hyperlinks-property-powerpoint%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/master-name-property-powerpoint%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/master-parent-property-powerpoint%28Office.15%29.aspx)|
|[Shapes](http://msdn.microsoft.com/library/master-shapes-property-powerpoint%28Office.15%29.aspx)|
|[SlideShowTransition](http://msdn.microsoft.com/library/master-slideshowtransition-property-powerpoint%28Office.15%29.aspx)|
|[TextStyles](http://msdn.microsoft.com/library/master-textstyles-property-powerpoint%28Office.15%29.aspx)|
|[Theme](http://msdn.microsoft.com/library/master-theme-property-powerpoint%28Office.15%29.aspx)|
|[TimeLine](http://msdn.microsoft.com/library/master-timeline-property-powerpoint%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/master-width-property-powerpoint%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/object-model-powerpoint-vba-reference%28Office.15%29.aspx)
