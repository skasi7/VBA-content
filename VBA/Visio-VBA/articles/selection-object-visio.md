---
title: Selection Object (Visio)
keywords: vis_sdr.chm10220
f1_keywords:
- vis_sdr.chm10220
ms.prod: VISIO
api_name:
- Visio.Selection
ms.assetid: e5734140-6dbe-7de8-9695-1a22fb4ac628
---


# Selection Object (Visio)

Represents a subset of  **Shape** objects for a page or master to which an operation can be applied.


## Remarks

To retrieve a  **Selection** object that corresponds to the set of shapes selected in a window, use the **Selection** property of a **Window** object.

The default property of a  **Selection** object is **Item**.

After you retrieve a  **Selection** object, you can add or remove shapes by using the **Select** method.

By default, the items reported by a  **Selection** object do not include subselected or superselected **Shape** objects. Use the **IterationMode** property to control whether subselected and superselected **Shape** objects are reported. You can determine whether an individual item is subselected or superselected by using the **ItemStatus** property.


## Methods



|**Name**|
|:-----|
|[AddToContainers](http://msdn.microsoft.com/library/selection-addtocontainers-method-visio%28Office.15%29.aspx)|
|[AddToGroup](http://msdn.microsoft.com/library/selection-addtogroup-method-visio%28Office.15%29.aspx)|
|[Align](http://msdn.microsoft.com/library/selection-align-method-visio%28Office.15%29.aspx)|
|[AutomaticLink](http://msdn.microsoft.com/library/selection-automaticlink-method-visio%28Office.15%29.aspx)|
|[AvoidPageBreaks](http://msdn.microsoft.com/library/selection-avoidpagebreaks-method-visio%28Office.15%29.aspx)|
|[BoundingBox](http://msdn.microsoft.com/library/selection-boundingbox-method-visio%28Office.15%29.aspx)|
|[BreakLinkToData](http://msdn.microsoft.com/library/selection-breaklinktodata-method-visio%28Office.15%29.aspx)|
|[BringForward](http://msdn.microsoft.com/library/selection-bringforward-method-visio%28Office.15%29.aspx)|
|[BringToFront](http://msdn.microsoft.com/library/selection-bringtofront-method-visio%28Office.15%29.aspx)|
|[Combine](http://msdn.microsoft.com/library/selection-combine-method-visio%28Office.15%29.aspx)|
|[ConnectShapes](http://msdn.microsoft.com/library/selection-connectshapes-method-visio%28Office.15%29.aspx)|
|[ConvertToGroup](http://msdn.microsoft.com/library/selection-converttogroup-method-visio%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/selection-copy-method-visio%28Office.15%29.aspx)|
|[Cut](http://msdn.microsoft.com/library/selection-cut-method-visio%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/selection-delete-method-visio%28Office.15%29.aspx)|
|[DeleteEx](http://msdn.microsoft.com/library/selection-deleteex-method-visio%28Office.15%29.aspx)|
|[DeselectAll](http://msdn.microsoft.com/library/selection-deselectall-method-visio%28Office.15%29.aspx)|
|[Distribute](http://msdn.microsoft.com/library/selection-distribute-method-visio%28Office.15%29.aspx)|
|[DrawRegion](http://msdn.microsoft.com/library/selection-drawregion-method-visio%28Office.15%29.aspx)|
|[Duplicate](http://msdn.microsoft.com/library/selection-duplicate-method-visio%28Office.15%29.aspx)|
|[Export](http://msdn.microsoft.com/library/selection-export-method-visio%28Office.15%29.aspx)|
|[FitCurve](http://msdn.microsoft.com/library/selection-fitcurve-method-visio%28Office.15%29.aspx)|
|[Flip](http://msdn.microsoft.com/library/selection-flip-method-visio%28Office.15%29.aspx)|
|[FlipHorizontal](http://msdn.microsoft.com/library/selection-fliphorizontal-method-visio%28Office.15%29.aspx)|
|[FlipVertical](http://msdn.microsoft.com/library/selection-flipvertical-method-visio%28Office.15%29.aspx)|
|[Fragment](http://msdn.microsoft.com/library/selection-fragment-method-visio%28Office.15%29.aspx)|
|[GetCallouts](http://msdn.microsoft.com/library/selection-getcallouts-method-visio%28Office.15%29.aspx)|
|[GetContainers](http://msdn.microsoft.com/library/selection-getcontainers-method-visio%28Office.15%29.aspx)|
|[GetIDs](http://msdn.microsoft.com/library/selection-getids-method-visio%28Office.15%29.aspx)|
|[Group](http://msdn.microsoft.com/library/selection-group-method-visio%28Office.15%29.aspx)|
|[Intersect](http://msdn.microsoft.com/library/selection-intersect-method-visio%28Office.15%29.aspx)|
|[Join](http://msdn.microsoft.com/library/selection-join-method-visio%28Office.15%29.aspx)|
|[Layout](http://msdn.microsoft.com/library/selection-layout-method-visio%28Office.15%29.aspx)|
|[LayoutChangeDirection](http://msdn.microsoft.com/library/selection-layoutchangedirection-method-visio%28Office.15%29.aspx)|
|[LayoutIncremental](http://msdn.microsoft.com/library/selection-layoutincremental-method-visio%28Office.15%29.aspx)|
|[LinkToData](http://msdn.microsoft.com/library/selection-linktodata-method-visio%28Office.15%29.aspx)|
|[MemberOfContainersIntersection](http://msdn.microsoft.com/library/selection-memberofcontainersintersection-method-visio%28Office.15%29.aspx)|
|[MemberOfContainersUnion](http://msdn.microsoft.com/library/selection-memberofcontainersunion-method-visio%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/selection-move-method-visio%28Office.15%29.aspx)|
|[MoveToSubprocess](http://msdn.microsoft.com/library/selection-movetosubprocess-method-visio%28Office.15%29.aspx)|
|[Offset](http://msdn.microsoft.com/library/selection-offset-method-visio%28Office.15%29.aspx)|
|[RemoveFromContainers](http://msdn.microsoft.com/library/selection-removefromcontainers-method-visio%28Office.15%29.aspx)|
|[RemoveFromGroup](http://msdn.microsoft.com/library/selection-removefromgroup-method-visio%28Office.15%29.aspx)|
|[ReplaceShape](http://msdn.microsoft.com/library/selection-replaceshape-method-visio%28Office.15%29.aspx)|
|[Resize](http://msdn.microsoft.com/library/selection-resize-method-visio%28Office.15%29.aspx)|
|[ReverseEnds](http://msdn.microsoft.com/library/selection-reverseends-method-visio%28Office.15%29.aspx)|
|[Rotate](http://msdn.microsoft.com/library/selection-rotate-method-visio%28Office.15%29.aspx)|
|[Rotate90](http://msdn.microsoft.com/library/selection-rotate90-method-visio%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/selection-select-method-visio%28Office.15%29.aspx)|
|[SelectAll](http://msdn.microsoft.com/library/selection-selectall-method-visio%28Office.15%29.aspx)|
|[SendBackward](http://msdn.microsoft.com/library/selection-sendbackward-method-visio%28Office.15%29.aspx)|
|[SendToBack](http://msdn.microsoft.com/library/selection-sendtoback-method-visio%28Office.15%29.aspx)|
|[SetContainerFormat](http://msdn.microsoft.com/library/selection-setcontainerformat-method-visio%28Office.15%29.aspx)|
|[SetQuickStyle](http://msdn.microsoft.com/library/selection-setquickstyle-method-visio%28Office.15%29.aspx)|
|[Subtract](http://msdn.microsoft.com/library/selection-subtract-method-visio%28Office.15%29.aspx)|
|[SwapEnds](http://msdn.microsoft.com/library/selection-swapends-method-visio%28Office.15%29.aspx)|
|[Trim](http://msdn.microsoft.com/library/selection-trim-method-visio%28Office.15%29.aspx)|
|[Ungroup](http://msdn.microsoft.com/library/selection-ungroup-method-visio%28Office.15%29.aspx)|
|[Union](http://msdn.microsoft.com/library/selection-union-method-visio%28Office.15%29.aspx)|
|[UpdateAlignmentBox](http://msdn.microsoft.com/library/selection-updatealignmentbox-method-visio%28Office.15%29.aspx)|
|[VisualBoundingBox](http://msdn.microsoft.com/library/selection-visualboundingbox-method-visio%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/selection-application-property-visio%28Office.15%29.aspx)|
|[ContainingMaster](http://msdn.microsoft.com/library/selection-containingmaster-property-visio%28Office.15%29.aspx)|
|[ContainingMasterID](http://msdn.microsoft.com/library/selection-containingmasterid-property-visio%28Office.15%29.aspx)|
|[ContainingPage](http://msdn.microsoft.com/library/selection-containingpage-property-visio%28Office.15%29.aspx)|
|[ContainingPageID](http://msdn.microsoft.com/library/selection-containingpageid-property-visio%28Office.15%29.aspx)|
|[ContainingShape](http://msdn.microsoft.com/library/selection-containingshape-property-visio%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/selection-count-property-visio%28Office.15%29.aspx)|
|[DataGraphic](http://msdn.microsoft.com/library/selection-datagraphic-property-visio%28Office.15%29.aspx)|
|[Document](http://msdn.microsoft.com/library/selection-document-property-visio%28Office.15%29.aspx)|
|[EventList](http://msdn.microsoft.com/library/selection-eventlist-property-visio%28Office.15%29.aspx)|
|[FillStyle](http://msdn.microsoft.com/library/selection-fillstyle-property-visio%28Office.15%29.aspx)|
|[FillStyleKeepFmt](http://msdn.microsoft.com/library/selection-fillstylekeepfmt-property-visio%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/selection-item-property-visio%28Office.15%29.aspx)|
|[ItemStatus](http://msdn.microsoft.com/library/selection-itemstatus-property-visio%28Office.15%29.aspx)|
|[IterationMode](http://msdn.microsoft.com/library/selection-iterationmode-property-visio%28Office.15%29.aspx)|
|[LineStyle](http://msdn.microsoft.com/library/selection-linestyle-property-visio%28Office.15%29.aspx)|
|[LineStyleKeepFmt](http://msdn.microsoft.com/library/selection-linestylekeepfmt-property-visio%28Office.15%29.aspx)|
|[ObjectType](http://msdn.microsoft.com/library/selection-objecttype-property-visio%28Office.15%29.aspx)|
|[PersistsEvents](http://msdn.microsoft.com/library/selection-persistsevents-property-visio%28Office.15%29.aspx)|
|[Picture](http://msdn.microsoft.com/library/selection-picture-property-visio%28Office.15%29.aspx)|
|[PrimaryItem](http://msdn.microsoft.com/library/selection-primaryitem-property-visio%28Office.15%29.aspx)|
|[SelectionForDragCopy](http://msdn.microsoft.com/library/selection-selectionfordragcopy-property-visio%28Office.15%29.aspx)|
|[Stat](http://msdn.microsoft.com/library/selection-stat-property-visio%28Office.15%29.aspx)|
|[Style](http://msdn.microsoft.com/library/selection-style-property-visio%28Office.15%29.aspx)|
|[StyleKeepFmt](http://msdn.microsoft.com/library/selection-stylekeepfmt-property-visio%28Office.15%29.aspx)|
|[TextStyle](http://msdn.microsoft.com/library/selection-textstyle-property-visio%28Office.15%29.aspx)|
|[TextStyleKeepFmt](http://msdn.microsoft.com/library/selection-textstylekeepfmt-property-visio%28Office.15%29.aspx)|

