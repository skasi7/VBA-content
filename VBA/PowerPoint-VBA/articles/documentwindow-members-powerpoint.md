---
title: DocumentWindow Members (PowerPoint)
ms.prod: POWERPOINT
ms.assetid: 414ea08d-db8e-70da-0fab-7a92942d2348
---


# DocumentWindow Members (PowerPoint)

Represents a document window. The  **DocumentWindow** object is a member of the **[DocumentWindows](documentwindows-object-powerpoint.md)** collection. The **DocumentWindows** collection contains all the open document windows.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|**[Activate](documentwindow-activate-method-powerpoint.md)**|Activates the specified object.|
|**[Close](documentwindow-close-method-powerpoint.md)**|Closes the specified document window.|
|**[ExpandSection](documentwindow-expandsection-method-powerpoint.md)**|Expands the section in the current  **DocumentWindow**.|
|**[FitToPage](documentwindow-fittopage-method-powerpoint.md)**|Adjusts the size of the specified document window to accommodate the information that's currently displayed.|
|**[IsSectionExpanded](documentwindow-issectionexpanded-method-powerpoint.md)**|Indicates whether the selected section is expanded in the  **DocumentWindow**.|
|**[LargeScroll](documentwindow-largescroll-method-powerpoint.md)**|Scrolls through the specified document window by pages.|
|**[NewWindow](documentwindow-newwindow-method-powerpoint.md)**|Opens a new window that contains the same document that is displayed in the specified window. Returns a  **[DocumentWindow](documentwindow-object-powerpoint.md)** object that represents the new window.|
|**[PointsToScreenPixelsX](documentwindow-pointstoscreenpixelsx-method-powerpoint.md)**|Converts a horizontal measurement from points to pixels. Used to return a horizontal screen location for a text frame or shape. Returns the converted measurement as a  **Single**.|
|**[PointsToScreenPixelsY](documentwindow-pointstoscreenpixelsy-method-powerpoint.md)**|Converts a vertical measurement from points to pixels. Used to return a vertical screen location for a text frame or shape. Returns the converted measurement as a  **Single**.|
|**[RangeFromPoint](documentwindow-rangefrompoint-method-powerpoint.md)**|Returns the  **Shape** object that is located at the point specified by the screen position coordinate pair. If no shape is located at the coordinate pair specified, then the method returns **Nothing**.|
|**[ScrollIntoView](documentwindow-scrollintoview-method-powerpoint.md)**|Scrolls the document window so that items within a specified rectangular area are displayed in the document window or pane.|
|**[SmallScroll](documentwindow-smallscroll-method-powerpoint.md)**|Scrolls through the specified document window by lines and columns.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|**[Active](documentwindow-active-property-powerpoint.md)**|Returns whether the specified pane or window is active. Read-only.|
|**[ActivePane](documentwindow-activepane-property-powerpoint.md)**|Returns a  **[Pane](pane-object-powerpoint.md)** object that represents the active pane in the document window. Read-only.|
|**[Application](documentwindow-application-property-powerpoint.md)**|Returns an  **[Application](application-object-powerpoint.md)** object that represents the creator of the specified object.|
|**[BlackAndWhite](documentwindow-blackandwhite-property-powerpoint.md)**|Determines whether the document window display is black and white. Read/write.|
|**[Caption](documentwindow-caption-property-powerpoint.md)**|Returns the text that appears in the title bar of the document window. Read-only.|
|**[Height](documentwindow-height-property-powerpoint.md)**|Returns or sets the height of the specified object, in points. Read/write.|
|**[Left](documentwindow-left-property-powerpoint.md)**|Returns or sets a  **Single** that represents the distance in points from the left edge of the document, application, and slide show windows to the left edge of the application window's client area. Setting this property to a very large positive or negative value may position the window completely off the desktop. Read/write.|
|**[Panes](documentwindow-panes-property-powerpoint.md)**|Returns a  **[Panes](panes-object-powerpoint.md)** collection that represents the panes in the document window. Read-only.|
|**[Parent](documentwindow-parent-property-powerpoint.md)**|Returns the parent object for the specified object.|
|**[Presentation](documentwindow-presentation-property-powerpoint.md)**|Returns a  **[Presentation](presentation-object-powerpoint.md)** object that represents the presentation in which the specified document window or slide show window was created. Read-only.|
|**[Selection](documentwindow-selection-property-powerpoint.md)**|Returns a  **[Selection](selection-object-powerpoint.md)** object that represents the selection in the specified document window. Read-only.|
|**[SplitHorizontal](documentwindow-splithorizontal-property-powerpoint.md)**|Returns or sets the percentage of the document window width that the outline pane occupies in normal view. Corresponds to the pane divider position between the slide and outline panes. Read/write.|
|**[SplitVertical](documentwindow-splitvertical-property-powerpoint.md)**|Returns or sets the percentage of the document window height that the slide pane occupies in normal view. Corresponds to the pane divider position between the slide and notes panes. Read/write.|
|**[Top](documentwindow-top-property-powerpoint.md)**|Returns or sets a  **Single** that represents the distance in points from the top edge of the document, application, and slide show window to the top edge of the application window's client area. Read/write.|
|**[View](documentwindow-view-property-powerpoint.md)**|Returns a  **[View](view-object-powerpoint.md)** object that represents the view in the specified document window. Read-only.|
|**[ViewType](documentwindow-viewtype-property-powerpoint.md)**|Returns or sets the type of the view contained in the specified document window. Read/write.|
|**[Width](documentwindow-width-property-powerpoint.md)**|Returns or sets the width of the specified object, in points. Read/write.|
|**[WindowState](documentwindow-windowstate-property-powerpoint.md)**|Returns or sets the state of the specified window. Read/write.|

