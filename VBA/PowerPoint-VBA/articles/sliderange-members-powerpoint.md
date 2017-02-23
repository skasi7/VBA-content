---
title: SlideRange Members (PowerPoint)
ms.prod: POWERPOINT
ms.assetid: f819c56d-96d5-836d-0d1f-49e505696f34
---


# SlideRange Members (PowerPoint)
A collection that represents a notes page or a slide range, which is a set of slides that can contain as little as a single slide or as much as all the slides in a presentation. 

A collection that represents a notes page or a slide range, which is a set of slides that can contain as little as a single slide or as much as all the slides in a presentation. 


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[ApplyTemplate](sliderange-applytemplate-method-powerpoint.md)|Applies a design template to the specified slide range.|
|[ApplyTemplate2](sliderange-applytemplate2-method-powerpoint.md)|Applies a design template and theme variant to the slide range.|
|[ApplyTheme](sliderange-applytheme-method-powerpoint.md)|Applies a theme or design template to the specified range of slides.|
|[ApplyThemeColorScheme](sliderange-applythemecolorscheme-method-powerpoint.md)|Applies a color scheme to the specified range of slides.|
|[Copy](sliderange-copy-method-powerpoint.md)|Copies the specified object to the Clipboard.|
|[Cut](sliderange-cut-method-powerpoint.md)|Deletes the specified object and places it on the Clipboard.|
|[Delete](sliderange-delete-method-powerpoint.md)|Deletes the specified  **SlideRange** object.|
|[Duplicate](sliderange-duplicate-method-powerpoint.md)|Creates a duplicate of the specified  **SlideRange** object, adds the new range of slides to the **Slides** collection immediately after the slide range specified originally, and then returns a **SlideRange** object that represents the duplicate slides.|
|[Export](sliderange-export-method-powerpoint.md)|Exports a range of slides, using the specified graphics filter, and saves the exported file under the specified file name.|
|[Item](sliderange-item-method-powerpoint.md)|Returns a single  **Slide** object from the specified **SlideRange** collection.|
|[MoveTo](sliderange-moveto-method-powerpoint.md)|Moves the specified object to a specific location within the same collection, renumbering all other items in the collection appropriately.|
|[MoveToSectionStart](sliderange-movetosectionstart-method-powerpoint.md)|Moves the current position to the start of the specified section in the  **SlideRange** object.|
|[PublishSlides](sliderange-publishslides-method-powerpoint.md)|Creates a Web presentation (in HTML format) from any loaded presentation. You can view the published presentation in a Web browser.|
|[Select](sliderange-select-method-powerpoint.md)|Selects the specified object.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](sliderange-application-property-powerpoint.md)|Returns an  **[Application](application-object-powerpoint.md)** object that represents the creator of the specified object.|
|[Background](sliderange-background-property-powerpoint.md)|Returns a  **[ShapeRange](shaperange-object-powerpoint.md)** object that represents the slide background.|
|[BackgroundStyle](sliderange-backgroundstyle-property-powerpoint.md)|Sets or returns the background style of the specified object. Read/write.|
|[ColorScheme](sliderange-colorscheme-property-powerpoint.md)|Returns or sets the  **[ColorScheme](colorscheme-object-powerpoint.md)** object that represents the scheme colors for the specified slide, slide range, or slide master. Read/write.|
|[Comments](sliderange-comments-property-powerpoint.md)|Returns a  **[Comments](comments-object-powerpoint.md)** object that represents a collection of comments. Read-only.|
|[Count](sliderange-count-property-powerpoint.md)|Returns the number of objects in the specified collection. Read-only.|
|[CustomerData](sliderange-customerdata-property-powerpoint.md)|Returns a  **[CustomerData](customerdata-object-powerpoint.md)** object. Read-only.|
|[CustomLayout](sliderange-customlayout-property-powerpoint.md)|Returns a  **[CustomLayout](customlayout-object-powerpoint.md)** object that represents the custom layout associated with the specified range of slides. Read-only.|
|[Design](sliderange-design-property-powerpoint.md)|Returns a  **Design** object representing a design.|
|[DisplayMasterShapes](sliderange-displaymastershapes-property-powerpoint.md)|Determines whether the specified range of slides displays the background objects on the slide master. Read/write.|
|[FollowMasterBackground](sliderange-followmasterbackground-property-powerpoint.md)|Determines whether the range of slides follows the slide master background. Read/write.|
|[HasNotesPage](sliderange-hasnotespage-property-powerpoint.md)|Indicates whether the selected  **SlideRange** has media that resides on a notes page. Read-only.|
|[HeadersFooters](sliderange-headersfooters-property-powerpoint.md)|Returns a  **[HeadersFooters](headersfooters-object-powerpoint.md)** collection that represents the header, footer, date and time, and slide number associated with the slide, slide master, or range of slides. Read-only.|
|[Hyperlinks](sliderange-hyperlinks-property-powerpoint.md)|Returns a  **[Hyperlinks](hyperlinks-object-powerpoint.md)** collection that represents all the hyperlinks on the specified slide. Read-only.|
|[Layout](sliderange-layout-property-powerpoint.md)|Returns or sets a  **PpSlideLayout** constant that represents the slide layout. Read/write.|
|[Master](sliderange-master-property-powerpoint.md)|Returns a  **[Master](master-object-powerpoint.md)** object that represents the slide master. Read-only.|
|[Name](sliderange-name-property-powerpoint.md)|When a slide is inserted into a presentation, Microsoft PowerPoint automatically assigns it a name in the form Slide _n_, where _n_ is an integer that represents the order in which the slide was created in the presentation. For example, the first slide inserted into a presentation is automatically named Slide1. If you copy a slide from one presentation to another, the slide loses the name it had in the first presentation and is automatically assigned a new name in the second presentation. A slide range must contain exactly one slide. Read/write.|
|[NotesPage](sliderange-notespage-property-powerpoint.md)|Returns a  **[SlideRange](sliderange-object-powerpoint.md)** object that represents the notes pages for the specified slide or range of slides. Read-only.|
|[Parent](sliderange-parent-property-powerpoint.md)|Returns the parent object for the specified object.|
|[PrintSteps](sliderange-printsteps-property-powerpoint.md)|Returns the number of slides you'd need to print to simulate the builds on the specified slide, slide master, or range of slides. Read-only.|
|[sectionIndex](sliderange-sectionindex-property-powerpoint.md)|Returns the index of the selected section in the  **SlideRange**. Read-only.|
|[Shapes](sliderange-shapes-property-powerpoint.md)|Returns a  **[Shapes](shapes-object-powerpoint.md)** collection that represents all the elements that have been placed or inserted on the specified slide, slide master, or range of slides. Read-only.|
|[SlideID](sliderange-slideid-property-powerpoint.md)|Returns a unique ID number for the specified slide. Read-only.|
|[SlideIndex](sliderange-slideindex-property-powerpoint.md)|Returns the index number of the specified slide within the  **Slides** collection. Read-only.|
|[SlideNumber](sliderange-slidenumber-property-powerpoint.md)|Returns the slide number. Read-only.|
|[SlideShowTransition](sliderange-slideshowtransition-property-powerpoint.md)|Returns a  **[SlideShowTransition](slideshowtransition-object-powerpoint.md)** object that represents the special effects for the specified slide transition. Read-only.|
|[Tags](sliderange-tags-property-powerpoint.md)|Returns a  **[Tags](tags-object-powerpoint.md)** object that represents the tags for the specified object. Read-only.|
|[ThemeColorScheme](sliderange-themecolorscheme-property-powerpoint.md)|Returns a  **ThemeColorScheme** object that represents the color scheme associated with the specified range of slides. Read-only.|
|[TimeLine](sliderange-timeline-property-powerpoint.md)|Returns a  **[TimeLine](timeline-object-powerpoint.md)** object that represents the animation timeline for the slide. Read-only.|

