---
title: Page Object (Visio)
keywords: vis_sdr.chm10190
f1_keywords:
- vis_sdr.chm10190
ms.prod: VISIO
api_name:
- Visio.Page
ms.assetid: 7a7f37ab-b448-eb70-b4f1-c185dfbd511e
---


# Page Object (Visio)

Represents a drawing page, which can be either a foreground page or a background page.


## Remarks

The default property of a  **Page** object is **Name**.

To retrieve the active page in an instance, use the  **ActivePage** property of an **Application** object.

The members of a  **Document** object's **Pages** collection represent the pages in that document. To retrieve a page's shapes, use the **Shapes** property of a **Page** object.


## Events



|**Name**|
|:-----|
|[AfterReplaceShapes](http://msdn.microsoft.com/library/page-afterreplaceshapes-event-visio%28Office.15%29.aspx)|
|[BeforePageDelete](http://msdn.microsoft.com/library/page-beforepagedelete-event-visio%28Office.15%29.aspx)|
|[BeforeReplaceShapes](http://msdn.microsoft.com/library/page-beforereplaceshapes-event-visio%28Office.15%29.aspx)|
|[BeforeSelectionDelete](http://msdn.microsoft.com/library/page-beforeselectiondelete-event-visio%28Office.15%29.aspx)|
|[BeforeShapeDelete](http://msdn.microsoft.com/library/page-beforeshapedelete-event-visio%28Office.15%29.aspx)|
|[BeforeShapeTextEdit](http://msdn.microsoft.com/library/page-beforeshapetextedit-event-visio%28Office.15%29.aspx)|
|[CalloutRelationshipAdded](http://msdn.microsoft.com/library/page-calloutrelationshipadded-event-visio%28Office.15%29.aspx)|
|[CalloutRelationshipDeleted](http://msdn.microsoft.com/library/page-calloutrelationshipdeleted-event-visio%28Office.15%29.aspx)|
|[CellChanged](http://msdn.microsoft.com/library/page-cellchanged-event-visio%28Office.15%29.aspx)|
|[ConnectionsAdded](http://msdn.microsoft.com/library/page-connectionsadded-event-visio%28Office.15%29.aspx)|
|[ConnectionsDeleted](http://msdn.microsoft.com/library/page-connectionsdeleted-event-visio%28Office.15%29.aspx)|
|[ContainerRelationshipAdded](http://msdn.microsoft.com/library/page-containerrelationshipadded-event-visio%28Office.15%29.aspx)|
|[ContainerRelationshipDeleted](http://msdn.microsoft.com/library/page-containerrelationshipdeleted-event-visio%28Office.15%29.aspx)|
|[ConvertToGroupCanceled](http://msdn.microsoft.com/library/page-converttogroupcanceled-event-visio%28Office.15%29.aspx)|
|[FormulaChanged](http://msdn.microsoft.com/library/page-formulachanged-event-visio%28Office.15%29.aspx)|
|[GroupCanceled](http://msdn.microsoft.com/library/page-groupcanceled-event-visio%28Office.15%29.aspx)|
|[PageChanged](http://msdn.microsoft.com/library/page-pagechanged-event-visio%28Office.15%29.aspx)|
|[PageDeleteCanceled](http://msdn.microsoft.com/library/page-pagedeletecanceled-event-visio%28Office.15%29.aspx)|
|[QueryCancelConvertToGroup](http://msdn.microsoft.com/library/page-querycancelconverttogroup-event-visio%28Office.15%29.aspx)|
|[QueryCancelGroup](http://msdn.microsoft.com/library/page-querycancelgroup-event-visio%28Office.15%29.aspx)|
|[QueryCancelPageDelete](http://msdn.microsoft.com/library/page-querycancelpagedelete-event-visio%28Office.15%29.aspx)|
|[QueryCancelReplaceShapes](http://msdn.microsoft.com/library/page-querycancelreplaceshapes-event-visio%28Office.15%29.aspx)|
|[QueryCancelSelectionDelete](http://msdn.microsoft.com/library/page-querycancelselectiondelete-event-visio%28Office.15%29.aspx)|
|[QueryCancelUngroup](http://msdn.microsoft.com/library/page-querycancelungroup-event-visio%28Office.15%29.aspx)|
|[ReplaceShapesCanceled](http://msdn.microsoft.com/library/page-replaceshapescanceled-event-visio%28Office.15%29.aspx)|
|[SelectionAdded](http://msdn.microsoft.com/library/page-selectionadded-event-visio%28Office.15%29.aspx)|
|[SelectionDeleteCanceled](http://msdn.microsoft.com/library/page-selectiondeletecanceled-event-visio%28Office.15%29.aspx)|
|[ShapeAdded](http://msdn.microsoft.com/library/page-shapeadded-event-visio%28Office.15%29.aspx)|
|[ShapeChanged](http://msdn.microsoft.com/library/page-shapechanged-event-visio%28Office.15%29.aspx)|
|[ShapeDataGraphicChanged](http://msdn.microsoft.com/library/page-shapedatagraphicchanged-event-visio%28Office.15%29.aspx)|
|[ShapeExitedTextEdit](http://msdn.microsoft.com/library/page-shapeexitedtextedit-event-visio%28Office.15%29.aspx)|
|[ShapeLinkAdded](http://msdn.microsoft.com/library/page-shapelinkadded-event-visio%28Office.15%29.aspx)|
|[ShapeLinkDeleted](http://msdn.microsoft.com/library/page-shapelinkdeleted-event-visio%28Office.15%29.aspx)|
|[ShapeParentChanged](http://msdn.microsoft.com/library/page-shapeparentchanged-event-visio%28Office.15%29.aspx)|
|[TextChanged](http://msdn.microsoft.com/library/page-textchanged-event-visio%28Office.15%29.aspx)|
|[UngroupCanceled](http://msdn.microsoft.com/library/page-ungroupcanceled-event-visio%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[AddGuide](http://msdn.microsoft.com/library/page-addguide-method-visio%28Office.15%29.aspx)|
|[AutoConnectMany](http://msdn.microsoft.com/library/page-autoconnectmany-method-visio%28Office.15%29.aspx)|
|[AutoSizeDrawing](http://msdn.microsoft.com/library/page-autosizedrawing-method-visio%28Office.15%29.aspx)|
|[AvoidPageBreaks](http://msdn.microsoft.com/library/page-avoidpagebreaks-method-visio%28Office.15%29.aspx)|
|[BoundingBox](http://msdn.microsoft.com/library/page-boundingbox-method-visio%28Office.15%29.aspx)|
|[CenterDrawing](http://msdn.microsoft.com/library/page-centerdrawing-method-visio%28Office.15%29.aspx)|
|[CreateSelection](http://msdn.microsoft.com/library/page-createselection-method-visio%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/page-delete-method-visio%28Office.15%29.aspx)|
|[DrawArcByThreePoints](http://msdn.microsoft.com/library/page-drawarcbythreepoints-method-visio%28Office.15%29.aspx)|
|[DrawBezier](http://msdn.microsoft.com/library/page-drawbezier-method-visio%28Office.15%29.aspx)|
|[DrawCircularArc](http://msdn.microsoft.com/library/page-drawcirculararc-method-visio%28Office.15%29.aspx)|
|[DrawLine](http://msdn.microsoft.com/library/page-drawline-method-visio%28Office.15%29.aspx)|
|[DrawNURBS](http://msdn.microsoft.com/library/page-drawnurbs-method-visio%28Office.15%29.aspx)|
|[DrawOval](http://msdn.microsoft.com/library/page-drawoval-method-visio%28Office.15%29.aspx)|
|[DrawPolyline](http://msdn.microsoft.com/library/page-drawpolyline-method-visio%28Office.15%29.aspx)|
|[DrawQuarterArc](http://msdn.microsoft.com/library/page-drawquarterarc-method-visio%28Office.15%29.aspx)|
|[DrawRectangle](http://msdn.microsoft.com/library/page-drawrectangle-method-visio%28Office.15%29.aspx)|
|[DrawSpline](http://msdn.microsoft.com/library/page-drawspline-method-visio%28Office.15%29.aspx)|
|[Drop](http://msdn.microsoft.com/library/page-drop-method-visio%28Office.15%29.aspx)|
|[DropCallout](http://msdn.microsoft.com/library/page-dropcallout-method-visio%28Office.15%29.aspx)|
|[DropConnected](http://msdn.microsoft.com/library/page-dropconnected-method-visio%28Office.15%29.aspx)|
|[DropContainer](http://msdn.microsoft.com/library/page-dropcontainer-method-visio%28Office.15%29.aspx)|
|[DropIntoList](http://msdn.microsoft.com/library/page-dropintolist-method-visio%28Office.15%29.aspx)|
|[DropLegend](http://msdn.microsoft.com/library/page-droplegend-method-visio%28Office.15%29.aspx)|
|[DropLinked](http://msdn.microsoft.com/library/page-droplinked-method-visio%28Office.15%29.aspx)|
|[DropMany](http://msdn.microsoft.com/library/page-dropmany-method-visio%28Office.15%29.aspx)|
|[DropManyLinkedU](http://msdn.microsoft.com/library/page-dropmanylinkedu-method-visio%28Office.15%29.aspx)|
|[DropManyU](http://msdn.microsoft.com/library/page-dropmanyu-method-visio%28Office.15%29.aspx)|
|[Duplicate](http://msdn.microsoft.com/library/page-duplicate-method-visio%28Office.15%29.aspx)|
|[Export](http://msdn.microsoft.com/library/page-export-method-visio%28Office.15%29.aspx)|
|[GetCallouts](http://msdn.microsoft.com/library/page-getcallouts-method-visio%28Office.15%29.aspx)|
|[GetContainers](http://msdn.microsoft.com/library/page-getcontainers-method-visio%28Office.15%29.aspx)|
|[GetFormulas](http://msdn.microsoft.com/library/page-getformulas-method-visio%28Office.15%29.aspx)|
|[GetFormulasU](http://msdn.microsoft.com/library/page-getformulasu-method-visio%28Office.15%29.aspx)|
|[GetResults](http://msdn.microsoft.com/library/page-getresults-method-visio%28Office.15%29.aspx)|
|[GetShapesLinkedToData](http://msdn.microsoft.com/library/page-getshapeslinkedtodata-method-visio%28Office.15%29.aspx)|
|[GetShapesLinkedToDataRow](http://msdn.microsoft.com/library/page-getshapeslinkedtodatarow-method-visio%28Office.15%29.aspx)|
|[GetTheme](http://msdn.microsoft.com/library/page-gettheme-method-visio%28Office.15%29.aspx)|
|[GetThemeVariant](http://msdn.microsoft.com/library/page-getthemevariant-method-visio%28Office.15%29.aspx)|
|[Import](http://msdn.microsoft.com/library/page-import-method-visio%28Office.15%29.aspx)|
|[InsertFromFile](http://msdn.microsoft.com/library/page-insertfromfile-method-visio%28Office.15%29.aspx)|
|[InsertObject](http://msdn.microsoft.com/library/page-insertobject-method-visio%28Office.15%29.aspx)|
|[Layout](http://msdn.microsoft.com/library/page-layout-method-visio%28Office.15%29.aspx)|
|[LayoutChangeDirection](http://msdn.microsoft.com/library/page-layoutchangedirection-method-visio%28Office.15%29.aspx)|
|[LayoutIncremental](http://msdn.microsoft.com/library/page-layoutincremental-method-visio%28Office.15%29.aspx)|
|[LinkShapesToDataRows](http://msdn.microsoft.com/library/page-linkshapestodatarows-method-visio%28Office.15%29.aspx)|
|[OpenDrawWindow](http://msdn.microsoft.com/library/page-opendrawwindow-method-visio%28Office.15%29.aspx)|
|[Paste](http://msdn.microsoft.com/library/page-paste-method-visio%28Office.15%29.aspx)|
|[PasteSpecial](http://msdn.microsoft.com/library/page-pastespecial-method-visio%28Office.15%29.aspx)|
|[PasteToLocation](http://msdn.microsoft.com/library/page-pastetolocation-method-visio%28Office.15%29.aspx)|
|[Print](http://msdn.microsoft.com/library/page-print-method-visio%28Office.15%29.aspx)|
|[PrintTile](http://msdn.microsoft.com/library/page-printtile-method-visio%28Office.15%29.aspx)|
|[ResizeToFitContents](http://msdn.microsoft.com/library/page-resizetofitcontents-method-visio%28Office.15%29.aspx)|
|[SetFormulas](http://msdn.microsoft.com/library/page-setformulas-method-visio%28Office.15%29.aspx)|
|[SetResults](http://msdn.microsoft.com/library/page-setresults-method-visio%28Office.15%29.aspx)|
|[SetTheme](http://msdn.microsoft.com/library/page-settheme-method-visio%28Office.15%29.aspx)|
|[SetThemeVariant](http://msdn.microsoft.com/library/page-setthemevariant-method-visio%28Office.15%29.aspx)|
|[ShapeIDsToUniqueIDs](http://msdn.microsoft.com/library/page-shapeidstouniqueids-method-visio%28Office.15%29.aspx)|
|[SplitConnector](http://msdn.microsoft.com/library/page-splitconnector-method-visio%28Office.15%29.aspx)|
|[UniqueIDsToShapeIDs](http://msdn.microsoft.com/library/page-uniqueidstoshapeids-method-visio%28Office.15%29.aspx)|
|[VisualBoundingBox](http://msdn.microsoft.com/library/page-visualboundingbox-method-visio%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/page-application-property-visio%28Office.15%29.aspx)|
|[AutoSize](http://msdn.microsoft.com/library/page-autosize-property-visio%28Office.15%29.aspx)|
|[Background](http://msdn.microsoft.com/library/page-background-property-visio%28Office.15%29.aspx)|
|[BackPage](http://msdn.microsoft.com/library/page-backpage-property-visio%28Office.15%29.aspx)|
|[Comments](http://msdn.microsoft.com/library/page-comments-property-visio%28Office.15%29.aspx)|
|[Connects](http://msdn.microsoft.com/library/page-connects-property-visio%28Office.15%29.aspx)|
|[Document](http://msdn.microsoft.com/library/page-document-property-visio%28Office.15%29.aspx)|
|[EventList](http://msdn.microsoft.com/library/page-eventlist-property-visio%28Office.15%29.aspx)|
|[ID](http://msdn.microsoft.com/library/page-id-property-visio%28Office.15%29.aspx)|
|[Index](http://msdn.microsoft.com/library/page-index-property-visio%28Office.15%29.aspx)|
|[Layers](http://msdn.microsoft.com/library/page-layers-property-visio%28Office.15%29.aspx)|
|[LayoutRoutePassive](http://msdn.microsoft.com/library/page-layoutroutepassive-property-visio%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/page-name-property-visio%28Office.15%29.aspx)|
|[NameU](http://msdn.microsoft.com/library/page-nameu-property-visio%28Office.15%29.aspx)|
|[ObjectType](http://msdn.microsoft.com/library/page-objecttype-property-visio%28Office.15%29.aspx)|
|[OLEObjects](http://msdn.microsoft.com/library/page-oleobjects-property-visio%28Office.15%29.aspx)|
|[OriginalPage](http://msdn.microsoft.com/library/page-originalpage-property-visio%28Office.15%29.aspx)|
|[PageSheet](http://msdn.microsoft.com/library/page-pagesheet-property-visio%28Office.15%29.aspx)|
|[PersistsEvents](http://msdn.microsoft.com/library/page-persistsevents-property-visio%28Office.15%29.aspx)|
|[Picture](http://msdn.microsoft.com/library/page-picture-property-visio%28Office.15%29.aspx)|
|[PrintTileCount](http://msdn.microsoft.com/library/page-printtilecount-property-visio%28Office.15%29.aspx)|
|[ReviewerID](http://msdn.microsoft.com/library/page-reviewerid-property-visio%28Office.15%29.aspx)|
|[ShapeComments](http://msdn.microsoft.com/library/page-shapecomments-property-visio%28Office.15%29.aspx)|
|[Shapes](http://msdn.microsoft.com/library/page-shapes-property-visio%28Office.15%29.aspx)|
|[SpatialSearch](http://msdn.microsoft.com/library/page-spatialsearch-property-visio%28Office.15%29.aspx)|
|[Stat](http://msdn.microsoft.com/library/page-stat-property-visio%28Office.15%29.aspx)|
|[ThemeColors](http://msdn.microsoft.com/library/page-themecolors-property-visio%28Office.15%29.aspx)|
|[ThemeEffects](http://msdn.microsoft.com/library/page-themeeffects-property-visio%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/page-type-property-visio%28Office.15%29.aspx)|

