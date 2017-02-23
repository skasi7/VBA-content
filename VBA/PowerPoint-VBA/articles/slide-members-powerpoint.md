---
title: Slide Members (PowerPoint)
ms.prod: POWERPOINT
ms.assetid: 3e34272b-615c-fa3f-4f0c-ceeba3c8f130
---


# Slide Members (PowerPoint)
Represents a slide. The  **[Slides](slides-object-powerpoint.md)** collection contains all the **Slide** objects in a presentation.

Represents a slide. The  **[Slides](slides-object-powerpoint.md)** collection contains all the **Slide** objects in a presentation.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[ApplyTemplate](slide-applytemplate-method-powerpoint.md)|Applies a design template to the specified slide.|
|[ApplyTemplate2](slide-applytemplate2-method-powerpoint.md)|Applies a design template and theme variant to the slide.|
|[ApplyTheme](slide-applytheme-method-powerpoint.md)|Applies a theme or design template to the specified slide.|
|[ApplyThemeColorScheme](slide-applythemecolorscheme-method-powerpoint.md)|Applies a color scheme to the specified slide.|
|[Copy](slide-copy-method-powerpoint.md)|Copies the specified object to the Clipboard.|
|[Cut](slide-cut-method-powerpoint.md)|Deletes the specified object and places it on the Clipboard.|
|[Delete](slide-delete-method-powerpoint.md)|Deletes the specified  **Slide** object.|
|[Duplicate](slide-duplicate-method-powerpoint.md)|Creates a duplicate of the specified  **Slide** object, adds the new slide to the **Slides** collection immediately after the slide specified originally, and then returns a **Slide** object that represents the duplicate slide.|
|[Export](slide-export-method-powerpoint.md)|Exports a slide, using the specified graphics filter, and saves the exported file under the specified file name.|
|[MoveTo](slide-moveto-method-powerpoint.md)|Moves the specified object to a specific location within the same collection, renumbering all other items in the collection appropriately.|
|[MoveToSectionStart](slide-movetosectionstart-method-powerpoint.md)|Moves the current slide to the start of the specified section.|
|[PublishSlides](slide-publishslides-method-powerpoint.md)|Publishes the specified slide to the specified location.|
|[Select](slide-select-method-powerpoint.md)|Selects the specified object.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](slide-application-property-powerpoint.md)|Returns an  **[Application](application-object-powerpoint.md)** object that represents the creator of the specified object.|
|[Background](slide-background-property-powerpoint.md)|Returns a  **[ShapeRange](shaperange-object-powerpoint.md)** object that represents the slide background.|
|[BackgroundStyle](slide-backgroundstyle-property-powerpoint.md)|Sets or returns the background style of the specified object. Read/write.|
|[ColorScheme](slide-colorscheme-property-powerpoint.md)|Returns or sets the  **[ColorScheme](colorscheme-object-powerpoint.md)** object that represents the scheme colors for the specified slide, slide range, or slide master. Read/write.|
|[Comments](slide-comments-property-powerpoint.md)|Returns a  **[Comments](comments-object-powerpoint.md)** object that represents a collection of comments. Read-only.|
|[CustomerData](slide-customerdata-property-powerpoint.md)|Returns a  **[CustomerData](customerdata-object-powerpoint.md)** object. Read-only.|
|[CustomLayout](slide-customlayout-property-powerpoint.md)|Returns a  **[CustomLayout](customlayout-object-powerpoint.md)** object that represents the custom layout associated with the specified slide. Read-only.|
|[Design](slide-design-property-powerpoint.md)|Returns a  **Design** object representing a design.|
|[DisplayMasterShapes](slide-displaymastershapes-property-powerpoint.md)|Determines whether the specified slide displays the background objects on the slide master. Read/write.|
|[FollowMasterBackground](slide-followmasterbackground-property-powerpoint.md)|Determines whether the slide follows the slide master background. Read/write.|
|[HasNotesPage](slide-hasnotespage-property-powerpoint.md)|Indicates whether the selected  **Slide** has media that resides on a notes page. Read-only.|
|[HeadersFooters](slide-headersfooters-property-powerpoint.md)|Returns a  **[HeadersFooters](headersfooters-object-powerpoint.md)** collection that represents the header, footer, date and time, and slide number associated with the slide, slide master, or range of slides. Read-only.|
|[Hyperlinks](slide-hyperlinks-property-powerpoint.md)|Returns a  **[Hyperlinks](hyperlinks-object-powerpoint.md)** collection that represents all the hyperlinks on the specified slide. Read-only.|
|[Layout](slide-layout-property-powerpoint.md)|Returns or sets a  **PpSlideLayout** constant that represents the slide layout. Read/write.|
|[Master](slide-master-property-powerpoint.md)|Returns a  **[Master](master-object-powerpoint.md)** object that represents the slide master. Read-only.|
|[Name](slide-name-property-powerpoint.md)|When a slide is inserted into a presentation, Microsoft PowerPoint automatically assigns it a name in the form Slide _n_, where _n_ is an integer that represents the order in which the slide was created in the presentation. For example, the first slide inserted into a presentation is automatically named Slide1. If you copy a slide from one presentation to another, the slide loses the name it had in the first presentation and is automatically assigned a new name in the second presentation. A slide range must contain exactly one slide. Read/write **String**.|
|[NotesPage](slide-notespage-property-powerpoint.md)|Returns a  **[SlideRange](sliderange-object-powerpoint.md)** object that represents the notes pages for the specified slide or range of slides. Read-only.|
|[Parent](slide-parent-property-powerpoint.md)|Returns the parent object for the specified object.|
|[PrintSteps](slide-printsteps-property-powerpoint.md)|Returns the number of slides you'd need to print to simulate the builds on the specified slide, slide master, or range of slides. Read-only.|
|[sectionIndex](slide-sectionindex-property-powerpoint.md)|Returns the index of the selected section in the  **Slide** range. Read-only.|
|[Shapes](slide-shapes-property-powerpoint.md)|Returns a  **[Shapes](shapes-object-powerpoint.md)** collection that represents all the elements that have been placed or inserted on the specified slide, slide master, or range of slides. Read-only.|
|[SlideID](slide-slideid-property-powerpoint.md)|Returns a unique ID number for the specified slide. Read-only.|
|[SlideIndex](slide-slideindex-property-powerpoint.md)|Returns the index number of the specified slide within the  **Slides** collection. Read-only.|
|[SlideNumber](slide-slidenumber-property-powerpoint.md)|Returns the slide number. Read-only.|
|[SlideShowTransition](slide-slideshowtransition-property-powerpoint.md)|Returns a  **[SlideShowTransition](slideshowtransition-object-powerpoint.md)** object that represents the special effects for the specified slide transition. Read-only.|
|[Tags](slide-tags-property-powerpoint.md)|Returns a  **[Tags](tags-object-powerpoint.md)** object that represents the tags for the specified object. Read-only.|
|[ThemeColorScheme](slide-themecolorscheme-property-powerpoint.md)|Returns a  **ThemeColorScheme** object that represents the color scheme associated with the specified slide. Read-only.|
|[TimeLine](slide-timeline-property-powerpoint.md)|Returns a  **[TimeLine](timeline-object-powerpoint.md)** object that represents the animation timeline for the slide. Read-only.|

