---
title: Window Object (Visio)
keywords: vis_sdr.chm10305
f1_keywords:
- vis_sdr.chm10305
ms.prod: VISIO
api_name:
- Visio.Window
ms.assetid: 5b49eb0f-07ea-00c7-52f1-2a3115a4b8ae
---


# Window Object (Visio)

Represents an open window in a Microsoft Visio instance.


## Remarks

The default property of a  **Window** object is **Application**.

To retrieve


- the active window in an instance of Visio, use the  **ActiveWindow** property of an **Application** object.
    
- a  **Page** object that represents the page shown in the window, use the **Page** property of a **Window** object.
    
- a  **Document** object that represents the document displayed in that window, use the **Document** property.
    
- a  **Selection** object that represents the shapes selected in that window, use the **Selection** property.
    

 **Note**  Beginning with Microsoft Visio 2002, the following methods of the  **Window** object are obsolete: **AddToGroup**, **Cut**, **Combine**, **Copy**, **Delete**, **Duplicate**, **Fragment**, **Group**, **Intersect**, **Join RemoveFromGroup**, **Subtract**, **Trim**, and **Union**. Existing solutions that invoke these methods will continue to work properly; however, new or rebuilt solutions should use these methods with the **Selection** object.

In addition, the  **Window** object's **Paste** method is now obsolete. Use the **Paste** or **PasteSpecial** method of the **Page**, **Master**, or **Shape** object. (Use the **Shape** object in the case of group shapes.)


## Events



|**Name**|
|:-----|
|[BeforeWindowClosed](http://msdn.microsoft.com/library/window-beforewindowclosed-event-visio%28Office.15%29.aspx)|
|[BeforeWindowPageTurn](http://msdn.microsoft.com/library/window-beforewindowpageturn-event-visio%28Office.15%29.aspx)|
|[BeforeWindowSelDelete](http://msdn.microsoft.com/library/window-beforewindowseldelete-event-visio%28Office.15%29.aspx)|
|[KeyDown](http://msdn.microsoft.com/library/window-keydown-event-visio%28Office.15%29.aspx)|
|[KeyPress](http://msdn.microsoft.com/library/window-keypress-event-visio%28Office.15%29.aspx)|
|[KeyUp](http://msdn.microsoft.com/library/window-keyup-event-visio%28Office.15%29.aspx)|
|[MouseDown](http://msdn.microsoft.com/library/window-mousedown-event-visio%28Office.15%29.aspx)|
|[MouseMove](http://msdn.microsoft.com/library/window-mousemove-event-visio%28Office.15%29.aspx)|
|[MouseUp](http://msdn.microsoft.com/library/window-mouseup-event-visio%28Office.15%29.aspx)|
|[OnKeystrokeMessageForAddon](http://msdn.microsoft.com/library/window-onkeystrokemessageforaddon-event-visio%28Office.15%29.aspx)|
|[QueryCancelWindowClose](http://msdn.microsoft.com/library/window-querycancelwindowclose-event-visio%28Office.15%29.aspx)|
|[SelectionChanged](http://msdn.microsoft.com/library/window-selectionchanged-event-visio%28Office.15%29.aspx)|
|[ViewChanged](http://msdn.microsoft.com/library/window-viewchanged-event-visio%28Office.15%29.aspx)|
|[WindowActivated](http://msdn.microsoft.com/library/window-windowactivated-event-visio%28Office.15%29.aspx)|
|[WindowChanged](http://msdn.microsoft.com/library/window-windowchanged-event-visio%28Office.15%29.aspx)|
|[WindowCloseCanceled](http://msdn.microsoft.com/library/window-windowclosecanceled-event-visio%28Office.15%29.aspx)|
|[WindowTurnedToPage](http://msdn.microsoft.com/library/window-windowturnedtopage-event-visio%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Activate](http://msdn.microsoft.com/library/window-activate-method-visio%28Office.15%29.aspx)|
|[CenterViewOnShape](http://msdn.microsoft.com/library/window-centerviewonshape-method-visio%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/window-close-method-visio%28Office.15%29.aspx)|
|[DeselectAll](http://msdn.microsoft.com/library/window-deselectall-method-visio%28Office.15%29.aspx)|
|[DockedStencils](http://msdn.microsoft.com/library/window-dockedstencils-method-visio%28Office.15%29.aspx)|
|[GetViewRect](http://msdn.microsoft.com/library/window-getviewrect-method-visio%28Office.15%29.aspx)|
|[GetWindowRect](http://msdn.microsoft.com/library/window-getwindowrect-method-visio%28Office.15%29.aspx)|
|[NewWindow](http://msdn.microsoft.com/library/window-newwindow-method-visio%28Office.15%29.aspx)|
|[Scroll](http://msdn.microsoft.com/library/window-scroll-method-visio%28Office.15%29.aspx)|
|[ScrollViewTo](http://msdn.microsoft.com/library/window-scrollviewto-method-visio%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/window-select-method-visio%28Office.15%29.aspx)|
|[SelectAll](http://msdn.microsoft.com/library/window-selectall-method-visio%28Office.15%29.aspx)|
|[SetViewRect](http://msdn.microsoft.com/library/window-setviewrect-method-visio%28Office.15%29.aspx)|
|[SetWindowRect](http://msdn.microsoft.com/library/window-setwindowrect-method-visio%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AllowEditing](http://msdn.microsoft.com/library/window-allowediting-property-visio%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/window-application-property-visio%28Office.15%29.aspx)|
|[BackgroundColor](http://msdn.microsoft.com/library/window-backgroundcolor-property-visio%28Office.15%29.aspx)|
|[BackgroundColorGradient](http://msdn.microsoft.com/library/window-backgroundcolorgradient-property-visio%28Office.15%29.aspx)|
|[Caption](http://msdn.microsoft.com/library/window-caption-property-visio%28Office.15%29.aspx)|
|[Document](http://msdn.microsoft.com/library/window-document-property-visio%28Office.15%29.aspx)|
|[EventList](http://msdn.microsoft.com/library/window-eventlist-property-visio%28Office.15%29.aspx)|
|[ID](http://msdn.microsoft.com/library/window-id-property-visio%28Office.15%29.aspx)|
|[Index](http://msdn.microsoft.com/library/window-index-property-visio%28Office.15%29.aspx)|
|[InPlace](http://msdn.microsoft.com/library/window-inplace-property-visio%28Office.15%29.aspx)|
|[IsEditingOLE](http://msdn.microsoft.com/library/window-iseditingole-property-visio%28Office.15%29.aspx)|
|[IsEditingText](http://msdn.microsoft.com/library/window-iseditingtext-property-visio%28Office.15%29.aspx)|
|[Master](http://msdn.microsoft.com/library/window-master-property-visio%28Office.15%29.aspx)|
|[MasterShortcut](http://msdn.microsoft.com/library/window-mastershortcut-property-visio%28Office.15%29.aspx)|
|[MergeCaption](http://msdn.microsoft.com/library/window-mergecaption-property-visio%28Office.15%29.aspx)|
|[MergeClass](http://msdn.microsoft.com/library/window-mergeclass-property-visio%28Office.15%29.aspx)|
|[MergeID](http://msdn.microsoft.com/library/window-mergeid-property-visio%28Office.15%29.aspx)|
|[MergePosition](http://msdn.microsoft.com/library/window-mergeposition-property-visio%28Office.15%29.aspx)|
|[ObjectType](http://msdn.microsoft.com/library/window-objecttype-property-visio%28Office.15%29.aspx)|
|[Page](http://msdn.microsoft.com/library/window-page-property-visio%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/window-parent-property-visio%28Office.15%29.aspx)|
|[ParentWindow](http://msdn.microsoft.com/library/window-parentwindow-property-visio%28Office.15%29.aspx)|
|[PersistsEvents](http://msdn.microsoft.com/library/window-persistsevents-property-visio%28Office.15%29.aspx)|
|[ReviewerMarkupVisible](http://msdn.microsoft.com/library/window-reviewermarkupvisible-property-visio%28Office.15%29.aspx)|
|[ScrollLock](http://msdn.microsoft.com/library/window-scrolllock-property-visio%28Office.15%29.aspx)|
|[SelectedCell](http://msdn.microsoft.com/library/window-selectedcell-property-visio%28Office.15%29.aspx)|
|[SelectedDataRecordset](http://msdn.microsoft.com/library/window-selecteddatarecordset-property-visio%28Office.15%29.aspx)|
|[SelectedDataRowID](http://msdn.microsoft.com/library/window-selecteddatarowid-property-visio%28Office.15%29.aspx)|
|[SelectedMasters](http://msdn.microsoft.com/library/window-selectedmasters-property-visio%28Office.15%29.aspx)|
|[SelectedText](http://msdn.microsoft.com/library/window-selectedtext-property-visio%28Office.15%29.aspx)|
|[SelectedValidationIssue](http://msdn.microsoft.com/library/window-selectedvalidationissue-property-visio%28Office.15%29.aspx)|
|[Selection](http://msdn.microsoft.com/library/window-selection-property-visio%28Office.15%29.aspx)|
|[SelectionForDragCopy](http://msdn.microsoft.com/library/window-selectionfordragcopy-property-visio%28Office.15%29.aspx)|
|[Shape](http://msdn.microsoft.com/library/window-shape-property-visio%28Office.15%29.aspx)|
|[ShowConnectPoints](http://msdn.microsoft.com/library/window-showconnectpoints-property-visio%28Office.15%29.aspx)|
|[ShowGrid](http://msdn.microsoft.com/library/window-showgrid-property-visio%28Office.15%29.aspx)|
|[ShowGuides](http://msdn.microsoft.com/library/window-showguides-property-visio%28Office.15%29.aspx)|
|[ShowPageBreaks](http://msdn.microsoft.com/library/window-showpagebreaks-property-visio%28Office.15%29.aspx)|
|[ShowPageOutline](http://msdn.microsoft.com/library/window-showpageoutline-property-visio%28Office.15%29.aspx)|
|[ShowPageTabs](http://msdn.microsoft.com/library/window-showpagetabs-property-visio%28Office.15%29.aspx)|
|[ShowRulers](http://msdn.microsoft.com/library/window-showrulers-property-visio%28Office.15%29.aspx)|
|[ShowScrollBars](http://msdn.microsoft.com/library/window-showscrollbars-property-visio%28Office.15%29.aspx)|
|[Stat](http://msdn.microsoft.com/library/window-stat-property-visio%28Office.15%29.aspx)|
|[SubType](http://msdn.microsoft.com/library/window-subtype-property-visio%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/window-type-property-visio%28Office.15%29.aspx)|
|[ViewFit](http://msdn.microsoft.com/library/window-viewfit-property-visio%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/window-visible-property-visio%28Office.15%29.aspx)|
|[WindowHandle32](http://msdn.microsoft.com/library/window-windowhandle32-property-visio%28Office.15%29.aspx)|
|[Windows](http://msdn.microsoft.com/library/window-windows-property-visio%28Office.15%29.aspx)|
|[WindowState](http://msdn.microsoft.com/library/window-windowstate-property-visio%28Office.15%29.aspx)|
|[Zoom](http://msdn.microsoft.com/library/window-zoom-property-visio%28Office.15%29.aspx)|
|[ZoomBehavior](http://msdn.microsoft.com/library/window-zoombehavior-property-visio%28Office.15%29.aspx)|
|[ZoomLock](http://msdn.microsoft.com/library/window-zoomlock-property-visio%28Office.15%29.aspx)|

