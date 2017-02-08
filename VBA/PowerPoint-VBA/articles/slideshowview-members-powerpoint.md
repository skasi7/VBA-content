---
title: SlideShowView Members (PowerPoint)
ms.prod: POWERPOINT
ms.assetid: fe2aacef-7324-4d07-55e9-0dffcdbb2a6c
---


# SlideShowView Members (PowerPoint)
Represents the view in a slide show window.

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[DrawLine](slideshowview-drawline-method-powerpoint.md)|Draws a line in the specified slide show view.|
|[EndNamedShow](slideshowview-endnamedshow-method-powerpoint.md)|Switches from running a custom, or named, slide show to running the entire presentation of which the custom show is a subset. When the slide show advances from the current slide, the next slide displayed will be the next one in the entire presentation, not the next one in the custom slide show.|
|[EraseDrawing](slideshowview-erasedrawing-method-powerpoint.md)|Removes lines drawn during a slide show by using either the  **[DrawLine](slideshowview-drawline-method-powerpoint.md)** method or the pen tool.|
|[Exit](slideshowview-exit-method-powerpoint.md)|Ends the specified slide show.|
|[First](slideshowview-first-method-powerpoint.md)|Sets the specified slide show view to display the first slide in the presentation.|
|[FirstAnimationIsAutomatic](slideshowview-firstanimationisautomatic-method-powerpoint.md)|Returns  **True** if the current slide has an initial animation that runs automatically.|
|[GetClickCount](slideshowview-getclickcount-method-powerpoint.md)|Returns the number of mouse clicks that are defined for a slide.|
|[GetClickIndex](slideshowview-getclickindex-method-powerpoint.md)|Returns the index number of the current mouse click for an animation that is actively playing on a slide or has just finished.|
|[GotoClick](slideshowview-gotoclick-method-powerpoint.md)|Plays an animation associated with a specified mouse click and any animations that follow on the slide.|
|[GotoNamedShow](slideshowview-gotonamedshow-method-powerpoint.md)|Switches to the specified custom, or named, slide show during another slide show. When the slide show advances from the current slide, the next slide displayed will be the next one in the specified custom slide show, not the next one in current slide show.|
|[GotoSlide](slideshowview-gotoslide-method-powerpoint.md)|Switches to the specified slide during a slide show. You can specify whether you want the animation effects to be rerun.|
|[Last](slideshowview-last-method-powerpoint.md)|Sets the specified slide show view to display the last slide in the presentation.|
|[Next](slideshowview-next-method-powerpoint.md)|Displays the slide immediately following the slide that's currently displayed. |
|[Player](slideshowview-player-method-powerpoint.md)|Allows access to playback controls for the associated view in the current window.|
|[Previous](slideshowview-previous-method-powerpoint.md)|Shows the slide immediately preceding the slide that's currently displayed. |
|[ResetSlideTime](slideshowview-resetslidetime-method-powerpoint.md)|Resets the elapsed time (represented by the  **[SlideElapsedTime](slideshowview-slideelapsedtime-property-powerpoint.md)** property) for the slide that's currently displayed to 0 (zero).|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[AcceleratorsEnabled](slideshowview-acceleratorsenabled-property-powerpoint.md)|Determines whether shortcut keys are enabled during a slide show. Read/write.|
|[AdvanceMode](slideshowview-advancemode-property-powerpoint.md)|Returns a value that indicates how the slide show in the specified view advances. Read-only.|
|[Application](slideshowview-application-property-powerpoint.md)|Returns an  **[Application](application-object-powerpoint.md)** object that represents the creator of the specified object.|
|[CurrentShowPosition](slideshowview-currentshowposition-property-powerpoint.md)|Returns the position of the current slide within the slide show that is showing in the specified view. Read-only.|
|[IsNamedShow](slideshowview-isnamedshow-property-powerpoint.md)|Determines whether a custom (named) slide show is displayed in the specified slide show view. Read-only.|
|[LaserPointerEnabled](slideshowview-laserpointerenabled-property-powerpoint.md)|Returns  **true** if the current slide show pointer is a laser pointer. This property is applicable only while the slide show is running. Read/write.|
|[LastSlideViewed](slideshowview-lastslideviewed-property-powerpoint.md)|Returns a  **[Slide](slide-object-powerpoint.md)** object that represents the slide viewed immediately before the current slide in the specified slide show view.|
|[MediaControlsHeight](slideshowview-mediacontrolsheight-property-powerpoint.md)|Returns the height of the media control bounding box. Read-only.|
|[MediaControlsLeft](slideshowview-mediacontrolsleft-property-powerpoint.md)|Returns the distance, in points, from the left edge of the media control bounding box to the left edge of the  **Slide**. Read-only.|
|[MediaControlsTop](slideshowview-mediacontrolstop-property-powerpoint.md)|Returns the distance, in points, from the top edge of the media control bounding box to the top edge of the  **Slide** object. Read-only.|
|[MediaControlsVisible](slideshowview-mediacontrolsvisible-property-powerpoint.md)|Indicates whether the media controls are visible. Read-only.|
|[MediaControlsWidth](slideshowview-mediacontrolswidth-property-powerpoint.md)|Returns the width, in points, of the media control bounding box. Read-only.|
|[Parent](slideshowview-parent-property-powerpoint.md)|Returns the parent object for the specified object.|
|[PointerColor](slideshowview-pointercolor-property-powerpoint.md)|Returns a  **ColorFormat** object that represents the pointer color for the specified presentation during one slide show. Read-only.|
|[PointerType](slideshowview-pointertype-property-powerpoint.md)|Returns or sets the type of pointer used in the slide show. Read/write.|
|[PresentationElapsedTime](slideshowview-presentationelapsedtime-property-powerpoint.md)|Returns the number of seconds that have elapsed since the beginning of the specified slide show. Read-only.|
|[Slide](slideshowview-slide-property-powerpoint.md)|Returns a  **[Slide](slide-object-powerpoint.md)** object that represents the slide that's currently displayed in the specified slide show window view. Read-only.|
|[SlideElapsedTime](slideshowview-slideelapsedtime-property-powerpoint.md)|Returns the number of seconds that the current slide has been displayed. Read/write.|
|[SlideShowName](slideshowview-slideshowname-property-powerpoint.md)|Returns the name of the custom slide show that's currently running in the specified slide show view. Read-only.|
|[State](slideshowview-state-property-powerpoint.md)|Returns or sets the state of the slide show. Read/write.|
|[Zoom](slideshowview-zoom-property-powerpoint.md)|Returns the zoom setting of the specified slide show window view as a percentage of normal size. Read-only.|

