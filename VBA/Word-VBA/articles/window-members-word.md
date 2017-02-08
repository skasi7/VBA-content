---
title: Window Members (Word)
ms.prod: WORD
ms.assetid: c0dec747-3695-4f96-ea25-05b6494aad7e
---


# Window Members (Word)
Represents a window. Many document characteristics, such as scroll bars and rulers, are actually properties of the window.

Represents a window. Many document characteristics, such as scroll bars and rulers, are actually properties of the window.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Activate](window-activate-method-word.md)|Activates the specified window.|
|[Close](window-close-method-word.md)|Closes the specified window.|
|[GetPoint](window-getpoint-method-word.md)|Returns the screen coordinates of the specified range or shape.|
|[LargeScroll](window-largescroll-method-word.md)|Scrolls a window or pane by the specified number of screens.|
|[NewWindow](window-newwindow-method-word.md)|Opens a new window with the same document as the specified window. Returns a  **Window** object.|
|[PageScroll](window-pagescroll-method-word.md)|Scrolls through the specified pane or window page by page.|
|[PrintOut](window-printout-method-word.md)|Prints all or part of the document displayed in the specified window.|
|[RangeFromPoint](window-rangefrompoint-method-word.md)|Returns the  **Range** or **Shape** object that is located at the point specified by the screen position coordinate pair.|
|[ScrollIntoView](window-scrollintoview-method-word.md)|Scrolls through the document window so the specified range or shape is displayed in the document window.|
|[SetFocus](window-setfocus-method-word.md)|Sets the focus of the specified document window to the body of an e-mail message.|
|[SmallScroll](window-smallscroll-method-word.md)|Scrolls a window or pane by the specified number of lines.|
|[ToggleRibbon](window-toggleribbon-method-word.md)|Shows or hides the ribbon.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Active](window-active-property-word.md)| **True** if the specified window is active. Read-only **Boolean** .|
|[ActivePane](window-activepane-property-word.md)|Returns a  **[Pane](pane-object-word.md)** object that represents the active pane for the specified window. Read-only.|
|[Application](window-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft OfficeWord application.|
|[Caption](window-caption-property-word.md)|Returns or sets the caption text for the window that is displayed in the title bar of the document or application window. Read/write  **String** .|
|[Creator](window-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[DisplayHorizontalScrollBar](window-displayhorizontalscrollbar-property-word.md)| **True** if a horizontal scroll bar is displayed for the specified window. Read/write **Boolean** .|
|[DisplayLeftScrollBar](window-displayleftscrollbar-property-word.md)| **True** if the vertical scroll bar appears on the left side of the document window. Read/write **Boolean** .|
|[DisplayRightRuler](window-displayrightruler-property-word.md)| **True** if the vertical ruler appears on the right side of the document window in print layout view. Read/write **Boolean** .|
|[DisplayRulers](window-displayrulers-property-word.md)| **True** if rulers are displayed for the specified window or pane. Read/write **Boolean** .|
|[DisplayScreenTips](window-displayscreentips-property-word.md)| **True** if comments, footnotes, endnotes, and hyperlinks are displayed as tips. Read/write **Boolean** .|
|[DisplayVerticalRuler](window-displayverticalruler-property-word.md)| **True** if a vertical ruler is displayed for the specified window or pane. Read/write **Boolean** .|
|[DisplayVerticalScrollBar](window-displayverticalscrollbar-property-word.md)| **True** if a vertical scroll bar is displayed for the specified window. Read/write **Boolean** .|
|[Document](window-document-property-word.md)|Returns a  **[Document](document-object-word.md)** object associated with the specified pane, window, or selection. Read-only.|
|[DocumentMap](window-documentmap-property-word.md)| **True** if the document map is visible. Read/write **Boolean** .|
|[EnvelopeVisible](window-envelopevisible-property-word.md)| **True** if the e-mail message header is visible in the document window. The default value is **False** . Read/write **Boolean** .|
|[Height](window-height-property-word.md)|Returns or sets the height of the window. Read/write Long.|
|[HorizontalPercentScrolled](window-horizontalpercentscrolled-property-word.md)|Returns or sets the horizontal scroll position as a percentage of the document width. Read/write  **Long** .|
|[Hwnd](window-hwnd-property-word.md)|Returns a  **Long** that indicates the window handle of the specified window. Read-only.|
|[IMEMode](window-imemode-property-word.md)|Returns or sets the default start-up mode for the Japanese Input Method Editor (IME). Read/write  **WdIMEMode** .|
|[Index](window-index-property-word.md)|Returns a  **Long** that represents the position of an item in a collection. Read-only.|
|[Left](window-left-property-word.md)|Returns or sets a  **Long** that represents the horizontal position of the specified window, measured in points. Read/write.|
|[Next](window-next-property-word.md)|Returns the next document window in the collection of open document windows. Read-only.|
|[Panes](window-panes-property-word.md)|Returns a  **[Panes](panes-object-word.md)** collection that represents all the window panes for the specified window.|
|[Parent](window-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **Window** object.|
|[Previous](window-previous-property-word.md)|Returns the previous document window in the collection open document windows. Read-only.|
|[Selection](window-selection-property-word.md)|Returns the  **Selection** object that represents a selected range or the insertion point. Read-only.|
|[ShowSourceDocuments](window-showsourcedocuments-property-word.md)|Returns or sets a  **[WdShowSourceDocuments](wdshowsourcedocuments-enumeration-word.md)** constant that represents how Microsoft Word displays source documents after a compare and merge process. Read/write.|
|[Split](window-split-property-word.md)| **True** if the window is split into multiple panes. Read/write **Boolean** .|
|[SplitVertical](window-splitvertical-property-word.md)|Returns or sets the vertical split percentage for the specified window. Read/write  **Long** .|
|[StyleAreaWidth](window-styleareawidth-property-word.md)|Returns or sets the width of the style area in points. Read/write  **Single** .|
|[Thumbnails](window-thumbnails-property-word.md)|Sets or returns a  **Boolean** that represents whether thumbnail images of the pages in a document are displayed along the left side of the Microsoft Word document window.|
|[Top](window-top-property-word.md)|Returns or sets the vertical position of the specified document window, in points. Read/write  **Long** .|
|[Type](window-type-property-word.md)|Returns the window type. Read-only  **[WdWindowType](wdwindowtype-enumeration-word.md)** .|
|[UsableHeight](window-usableheight-property-word.md)|Returns the height (in points) of the active working area in the specified document window. Read-only  **Long** . .|
|[UsableWidth](window-usablewidth-property-word.md)|Returns the width (in points) of the active working area in the specified document window. Read-only  **Long** .|
|[VerticalPercentScrolled](window-verticalpercentscrolled-property-word.md)|Returns or sets the vertical scroll position as a percentage of the document length. Read/write  **Long** .|
|[View](window-view-property-word.md)|Returns a  **[View](view-object-word.md)** object that represents the view for the specified window or pane.|
|[Visible](window-visible-property-word.md)| **True** if the specified object is visible. Read/write **Boolean** .|
|[Width](window-width-property-word.md)|Returns or sets the width of the specified document window, in points. Read/write  **Long** .|
|[WindowNumber](window-windownumber-property-word.md)|Returns the window number of the document displayed in the specified window. For example, if the caption of the window is "Sales.doc:2", this property returns the number 2. Read-only  **Long** .|
|[WindowState](window-windowstate-property-word.md)|Returns or sets the state of the specified document window or task window. Read/write  **[WdWindowState](wdwindowstate-enumeration-word.md)** .|

