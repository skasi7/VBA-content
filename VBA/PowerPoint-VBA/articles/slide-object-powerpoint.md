---
title: Slide Object (PowerPoint)
keywords: vbapp10.chm535000
f1_keywords:
- vbapp10.chm535000
ms.prod: POWERPOINT
ms.assetid: afe42344-6898-00d2-ecc1-b0ed23a71fe8
---


# Slide Object (PowerPoint)

Represents a slide. The  **[Slides](http://msdn.microsoft.com/library/slides-object-powerpoint%28Office.15%29.aspx)** collection contains all the **Slide** objects in a presentation.


## Remarks


 **Note**  Don't be confused if you're trying to return a reference to a single slide but you end up with a  **[SlideRange](http://msdn.microsoft.com/library/sliderange-object-powerpoint%28Office.15%29.aspx)** object. A single slide can be represented either by a **Slide** object or by a[SlideRange](http://msdn.microsoft.com/library/sliderange-object-powerpoint%28Office.15%29.aspx)collection that contains only one slide, depending on how you return a reference to the slide. For example, if you create and return a reference to a slide by using the  **[Add](http://msdn.microsoft.com/library/presentations-add-method-powerpoint%28Office.15%29.aspx)** method, the slide is represented by a **Slide** object. However, if you create and return a reference to a slide by using the **[Duplicate](http://msdn.microsoft.com/library/slide-duplicate-method-powerpoint%28Office.15%29.aspx)** method, the slide is represented by a **SlideRange** collection that contains a single slide. Because all the properties and methods that apply to a **Slide** object also apply to a **SlideRange** collection that contains a single slide, you can work with the returned slide in the same way, regardless of whether it is represented by a **Slide** object or a **SlideRange** collection.

The following examples describe how to:


- Return a slide that you specify by name, index number, or slide ID number
    
- Return a slide in the selection
    
- Return the slide that's currently displayed in any document window or slide show window you specify
    
- Create a new slide
    

## Example

Use  **Slides** (index), where index is the slide name or index number, or use **Slides.FindBySlideID** (index), where index is the slide ID number, to return a single **Slide** object. The following example sets the layout for slide one in the active presentation.


```
ActivePresentation.Slides(1).Layout = ppLayoutTitle
```

The following example sets the layout for the slide with the ID number 265.




```
ActivePresentation.Slides.FindBySlideID(265).Layout = ppLayoutTitle
```

Use  **Selection.SlideRange** (index), where index is the slide name or index number within the selection, to return a single **Slide** object. The following example sets the layout for slide one in the selection in the active window, assuming that there's at least one slide selected.




```
ActiveWindow.Selection.SlideRange(1).Layout = ppLayoutTitle
```

If there's only one slide selected, you can use  **Selection.SlideRange** to return a **SlideRange** collection that contains the selected slide. The following example sets the layout for slide one in the current selection in the active window, assuming that there's exactly one slide selected.




```
ActiveWindow.Selection.SlideRange.Layout = ppLayoutTitle
```

Use the  **Slide** property to return the slide that's currently displayed in the specified document window or slide show window view. The following example copies the slide that's currently displayed in document window two to the Clipboard.




```
Windows(2).View.Slide.Copy
```

Use the  **Add** method to create a new slide and add it to the presentation. The following example adds a title slide to the beginning of the active presentation.




```
ActivePresentation.Slides.Add 1, ppLayoutTitleOnly
```


## Methods



|**Name**|
|:-----|
|[ApplyTemplate](http://msdn.microsoft.com/library/slide-applytemplate-method-powerpoint%28Office.15%29.aspx)|
|[ApplyTemplate2](http://msdn.microsoft.com/library/slide-applytemplate2-method-powerpoint%28Office.15%29.aspx)|
|[ApplyTheme](http://msdn.microsoft.com/library/slide-applytheme-method-powerpoint%28Office.15%29.aspx)|
|[ApplyThemeColorScheme](http://msdn.microsoft.com/library/slide-applythemecolorscheme-method-powerpoint%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/slide-copy-method-powerpoint%28Office.15%29.aspx)|
|[Cut](http://msdn.microsoft.com/library/slide-cut-method-powerpoint%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/slide-delete-method-powerpoint%28Office.15%29.aspx)|
|[Duplicate](http://msdn.microsoft.com/library/slide-duplicate-method-powerpoint%28Office.15%29.aspx)|
|[Export](http://msdn.microsoft.com/library/slide-export-method-powerpoint%28Office.15%29.aspx)|
|[MoveTo](http://msdn.microsoft.com/library/slide-moveto-method-powerpoint%28Office.15%29.aspx)|
|[MoveToSectionStart](http://msdn.microsoft.com/library/slide-movetosectionstart-method-powerpoint%28Office.15%29.aspx)|
|[PublishSlides](http://msdn.microsoft.com/library/slide-publishslides-method-powerpoint%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/slide-select-method-powerpoint%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/slide-application-property-powerpoint%28Office.15%29.aspx)|
|[Background](http://msdn.microsoft.com/library/slide-background-property-powerpoint%28Office.15%29.aspx)|
|[BackgroundStyle](http://msdn.microsoft.com/library/slide-backgroundstyle-property-powerpoint%28Office.15%29.aspx)|
|[ColorScheme](http://msdn.microsoft.com/library/slide-colorscheme-property-powerpoint%28Office.15%29.aspx)|
|[Comments](http://msdn.microsoft.com/library/slide-comments-property-powerpoint%28Office.15%29.aspx)|
|[CustomerData](http://msdn.microsoft.com/library/slide-customerdata-property-powerpoint%28Office.15%29.aspx)|
|[CustomLayout](http://msdn.microsoft.com/library/slide-customlayout-property-powerpoint%28Office.15%29.aspx)|
|[Design](http://msdn.microsoft.com/library/slide-design-property-powerpoint%28Office.15%29.aspx)|
|[DisplayMasterShapes](http://msdn.microsoft.com/library/slide-displaymastershapes-property-powerpoint%28Office.15%29.aspx)|
|[FollowMasterBackground](http://msdn.microsoft.com/library/slide-followmasterbackground-property-powerpoint%28Office.15%29.aspx)|
|[HasNotesPage](http://msdn.microsoft.com/library/slide-hasnotespage-property-powerpoint%28Office.15%29.aspx)|
|[HeadersFooters](http://msdn.microsoft.com/library/slide-headersfooters-property-powerpoint%28Office.15%29.aspx)|
|[Hyperlinks](http://msdn.microsoft.com/library/slide-hyperlinks-property-powerpoint%28Office.15%29.aspx)|
|[Layout](http://msdn.microsoft.com/library/slide-layout-property-powerpoint%28Office.15%29.aspx)|
|[Master](http://msdn.microsoft.com/library/slide-master-property-powerpoint%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/slide-name-property-powerpoint%28Office.15%29.aspx)|
|[NotesPage](http://msdn.microsoft.com/library/slide-notespage-property-powerpoint%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/slide-parent-property-powerpoint%28Office.15%29.aspx)|
|[PrintSteps](http://msdn.microsoft.com/library/slide-printsteps-property-powerpoint%28Office.15%29.aspx)|
|[sectionIndex](http://msdn.microsoft.com/library/slide-sectionindex-property-powerpoint%28Office.15%29.aspx)|
|[Shapes](http://msdn.microsoft.com/library/slide-shapes-property-powerpoint%28Office.15%29.aspx)|
|[SlideID](http://msdn.microsoft.com/library/slide-slideid-property-powerpoint%28Office.15%29.aspx)|
|[SlideIndex](http://msdn.microsoft.com/library/slide-slideindex-property-powerpoint%28Office.15%29.aspx)|
|[SlideNumber](http://msdn.microsoft.com/library/slide-slidenumber-property-powerpoint%28Office.15%29.aspx)|
|[SlideShowTransition](http://msdn.microsoft.com/library/slide-slideshowtransition-property-powerpoint%28Office.15%29.aspx)|
|[Tags](http://msdn.microsoft.com/library/slide-tags-property-powerpoint%28Office.15%29.aspx)|
|[ThemeColorScheme](http://msdn.microsoft.com/library/slide-themecolorscheme-property-powerpoint%28Office.15%29.aspx)|
|[TimeLine](http://msdn.microsoft.com/library/slide-timeline-property-powerpoint%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/object-model-powerpoint-vba-reference%28Office.15%29.aspx)
