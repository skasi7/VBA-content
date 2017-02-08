---
title: Shape Object (Visio)
keywords: vis_sdr.chm10225
f1_keywords:
- vis_sdr.chm10225
ms.prod: VISIO
api_name:
- Visio.Shape
ms.assetid: da7a8872-4ebb-a607-e0ed-eebf68ff5630
---


# Shape Object (Visio)

Represents anything you can select in a drawing window: a basic shape, a group, a guide, or an object from another application embedded or linked in Microsoft Visio.


## Remarks

The default property of a  **Shape** object is **Name**.

You can retrieve a particular  **Shape** object from the **Shapes** collection of the following objects:




-  **Page** object
    
-  **Master** object
    
-  **Shape** object that represents a group
    


To retrieve  **Cell** objects and **Connect** objects, use the **Cells** and **Connects** properties of a **Shape** object, respectively.


 **Note**  The **PageSheet** property of a **Page** object and **Master** object returns a **Shape** object whose **Type** property returns **visTypePage**. It has cells that specify properties such as drawing size and drawing scale. The **DocumentSheet** property of a **Document** object also returns a **Shape** object whose **Type** property returns **visTypeDoc**. It has cells that specify properties of the document.If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this object maps to the following types:


## Events



|**Name**|
|:-----|
|[BeforeSelectionDelete](http://msdn.microsoft.com/library/shape-beforeselectiondelete-event-visio%28Office.15%29.aspx)|
|[BeforeShapeDelete](http://msdn.microsoft.com/library/shape-beforeshapedelete-event-visio%28Office.15%29.aspx)|
|[BeforeShapeTextEdit](http://msdn.microsoft.com/library/shape-beforeshapetextedit-event-visio%28Office.15%29.aspx)|
|[CellChanged](http://msdn.microsoft.com/library/shape-cellchanged-event-visio%28Office.15%29.aspx)|
|[ConvertToGroupCanceled](http://msdn.microsoft.com/library/shape-converttogroupcanceled-event-visio%28Office.15%29.aspx)|
|[FormulaChanged](http://msdn.microsoft.com/library/shape-formulachanged-event-visio%28Office.15%29.aspx)|
|[GroupCanceled](http://msdn.microsoft.com/library/shape-groupcanceled-event-visio%28Office.15%29.aspx)|
|[QueryCancelConvertToGroup](http://msdn.microsoft.com/library/shape-querycancelconverttogroup-event-visio%28Office.15%29.aspx)|
|[QueryCancelGroup](http://msdn.microsoft.com/library/shape-querycancelgroup-event-visio%28Office.15%29.aspx)|
|[QueryCancelSelectionDelete](http://msdn.microsoft.com/library/shape-querycancelselectiondelete-event-visio%28Office.15%29.aspx)|
|[QueryCancelUngroup](http://msdn.microsoft.com/library/shape-querycancelungroup-event-visio%28Office.15%29.aspx)|
|[SelectionAdded](http://msdn.microsoft.com/library/shape-selectionadded-event-visio%28Office.15%29.aspx)|
|[SelectionDeleteCanceled](http://msdn.microsoft.com/library/shape-selectiondeletecanceled-event-visio%28Office.15%29.aspx)|
|[ShapeAdded](http://msdn.microsoft.com/library/shape-shapeadded-event-visio%28Office.15%29.aspx)|
|[ShapeChanged](http://msdn.microsoft.com/library/shape-shapechanged-event-visio%28Office.15%29.aspx)|
|[ShapeDataGraphicChanged](http://msdn.microsoft.com/library/shape-shapedatagraphicchanged-event-visio%28Office.15%29.aspx)|
|[ShapeExitedTextEdit](http://msdn.microsoft.com/library/shape-shapeexitedtextedit-event-visio%28Office.15%29.aspx)|
|[ShapeLinkAdded](http://msdn.microsoft.com/library/shape-shapelinkadded-event-visio%28Office.15%29.aspx)|
|[ShapeLinkDeleted](http://msdn.microsoft.com/library/shape-shapelinkdeleted-event-visio%28Office.15%29.aspx)|
|[ShapeParentChanged](http://msdn.microsoft.com/library/shape-shapeparentchanged-event-visio%28Office.15%29.aspx)|
|[TextChanged](http://msdn.microsoft.com/library/shape-textchanged-event-visio%28Office.15%29.aspx)|
|[UngroupCanceled](http://msdn.microsoft.com/library/shape-ungroupcanceled-event-visio%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[AddGuide](http://msdn.microsoft.com/library/shape-addguide-method-visio%28Office.15%29.aspx)|
|[AddHyperlink](http://msdn.microsoft.com/library/shape-addhyperlink-method-visio%28Office.15%29.aspx)|
|[AddNamedRow](http://msdn.microsoft.com/library/shape-addnamedrow-method-visio%28Office.15%29.aspx)|
|[AddRow](http://msdn.microsoft.com/library/shape-addrow-method-visio%28Office.15%29.aspx)|
|[AddRows](http://msdn.microsoft.com/library/shape-addrows-method-visio%28Office.15%29.aspx)|
|[AddSection](http://msdn.microsoft.com/library/shape-addsection-method-visio%28Office.15%29.aspx)|
|[AddToContainers](http://msdn.microsoft.com/library/shape-addtocontainers-method-visio%28Office.15%29.aspx)|
|[AutoConnect](http://msdn.microsoft.com/library/shape-autoconnect-method-visio%28Office.15%29.aspx)|
|[BoundingBox](http://msdn.microsoft.com/library/shape-boundingbox-method-visio%28Office.15%29.aspx)|
|[BreakLinkToData](http://msdn.microsoft.com/library/shape-breaklinktodata-method-visio%28Office.15%29.aspx)|
|[BringForward](http://msdn.microsoft.com/library/shape-bringforward-method-visio%28Office.15%29.aspx)|
|[BringToFront](http://msdn.microsoft.com/library/shape-bringtofront-method-visio%28Office.15%29.aspx)|
|[CenterDrawing](http://msdn.microsoft.com/library/shape-centerdrawing-method-visio%28Office.15%29.aspx)|
|[ChangePicture](http://msdn.microsoft.com/library/shape-changepicture-method-visio%28Office.15%29.aspx)|
|[ConnectedShapes](http://msdn.microsoft.com/library/shape-connectedshapes-method-visio%28Office.15%29.aspx)|
|[ConvertToGroup](http://msdn.microsoft.com/library/shape-converttogroup-method-visio%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/shape-copy-method-visio%28Office.15%29.aspx)|
|[CreateSelection](http://msdn.microsoft.com/library/shape-createselection-method-visio%28Office.15%29.aspx)|
|[CreateSubProcess](http://msdn.microsoft.com/library/shape-createsubprocess-method-visio%28Office.15%29.aspx)|
|[Cut](http://msdn.microsoft.com/library/shape-cut-method-visio%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/shape-delete-method-visio%28Office.15%29.aspx)|
|[DeleteEx](http://msdn.microsoft.com/library/shape-deleteex-method-visio%28Office.15%29.aspx)|
|[DeleteRow](http://msdn.microsoft.com/library/shape-deleterow-method-visio%28Office.15%29.aspx)|
|[DeleteSection](http://msdn.microsoft.com/library/shape-deletesection-method-visio%28Office.15%29.aspx)|
|[Disconnect](http://msdn.microsoft.com/library/shape-disconnect-method-visio%28Office.15%29.aspx)|
|[DrawArcByThreePoints](http://msdn.microsoft.com/library/shape-drawarcbythreepoints-method-visio%28Office.15%29.aspx)|
|[DrawBezier](http://msdn.microsoft.com/library/shape-drawbezier-method-visio%28Office.15%29.aspx)|
|[DrawCircularArc](http://msdn.microsoft.com/library/shape-drawcirculararc-method-visio%28Office.15%29.aspx)|
|[DrawLine](http://msdn.microsoft.com/library/shape-drawline-method-visio%28Office.15%29.aspx)|
|[DrawNURBS](http://msdn.microsoft.com/library/shape-drawnurbs-method-visio%28Office.15%29.aspx)|
|[DrawOval](http://msdn.microsoft.com/library/shape-drawoval-method-visio%28Office.15%29.aspx)|
|[DrawPolyline](http://msdn.microsoft.com/library/shape-drawpolyline-method-visio%28Office.15%29.aspx)|
|[DrawQuarterArc](http://msdn.microsoft.com/library/shape-drawquarterarc-method-visio%28Office.15%29.aspx)|
|[DrawRectangle](http://msdn.microsoft.com/library/shape-drawrectangle-method-visio%28Office.15%29.aspx)|
|[DrawSpline](http://msdn.microsoft.com/library/shape-drawspline-method-visio%28Office.15%29.aspx)|
|[Drop](http://msdn.microsoft.com/library/shape-drop-method-visio%28Office.15%29.aspx)|
|[DropMany](http://msdn.microsoft.com/library/shape-dropmany-method-visio%28Office.15%29.aspx)|
|[DropManyU](http://msdn.microsoft.com/library/shape-dropmanyu-method-visio%28Office.15%29.aspx)|
|[Duplicate](http://msdn.microsoft.com/library/shape-duplicate-method-visio%28Office.15%29.aspx)|
|[Export](http://msdn.microsoft.com/library/shape-export-method-visio%28Office.15%29.aspx)|
|[FitCurve](http://msdn.microsoft.com/library/shape-fitcurve-method-visio%28Office.15%29.aspx)|
|[FlipHorizontal](http://msdn.microsoft.com/library/shape-fliphorizontal-method-visio%28Office.15%29.aspx)|
|[FlipVertical](http://msdn.microsoft.com/library/shape-flipvertical-method-visio%28Office.15%29.aspx)|
|[GetCustomPropertiesLinkedToData](http://msdn.microsoft.com/library/shape-getcustompropertieslinkedtodata-method-visio%28Office.15%29.aspx)|
|[GetCustomPropertyLinkedColumn](http://msdn.microsoft.com/library/shape-getcustompropertylinkedcolumn-method-visio%28Office.15%29.aspx)|
|[GetFormulas](http://msdn.microsoft.com/library/shape-getformulas-method-visio%28Office.15%29.aspx)|
|[GetFormulasU](http://msdn.microsoft.com/library/shape-getformulasu-method-visio%28Office.15%29.aspx)|
|[GetLinkedDataRecordsetIDs](http://msdn.microsoft.com/library/shape-getlinkeddatarecordsetids-method-visio%28Office.15%29.aspx)|
|[GetLinkedDataRow](http://msdn.microsoft.com/library/shape-getlinkeddatarow-method-visio%28Office.15%29.aspx)|
|[GetResults](http://msdn.microsoft.com/library/shape-getresults-method-visio%28Office.15%29.aspx)|
|[GluedShapes](http://msdn.microsoft.com/library/shape-gluedshapes-method-visio%28Office.15%29.aspx)|
|[Group](http://msdn.microsoft.com/library/shape-group-method-visio%28Office.15%29.aspx)|
|[HasCategory](http://msdn.microsoft.com/library/shape-hascategory-method-visio%28Office.15%29.aspx)|
|[HitTest](http://msdn.microsoft.com/library/shape-hittest-method-visio%28Office.15%29.aspx)|
|[Import](http://msdn.microsoft.com/library/shape-import-method-visio%28Office.15%29.aspx)|
|[InsertFromFile](http://msdn.microsoft.com/library/shape-insertfromfile-method-visio%28Office.15%29.aspx)|
|[InsertObject](http://msdn.microsoft.com/library/shape-insertobject-method-visio%28Office.15%29.aspx)|
|[IsCustomPropertyLinked](http://msdn.microsoft.com/library/shape-iscustompropertylinked-method-visio%28Office.15%29.aspx)|
|[Layout](http://msdn.microsoft.com/library/shape-layout-method-visio%28Office.15%29.aspx)|
|[LinkToData](http://msdn.microsoft.com/library/shape-linktodata-method-visio%28Office.15%29.aspx)|
|[MoveToSubprocess](http://msdn.microsoft.com/library/shape-movetosubprocess-method-visio%28Office.15%29.aspx)|
|[Offset](http://msdn.microsoft.com/library/shape-offset-method-visio%28Office.15%29.aspx)|
|[OpenDrawWindow](http://msdn.microsoft.com/library/shape-opendrawwindow-method-visio%28Office.15%29.aspx)|
|[OpenSheetWindow](http://msdn.microsoft.com/library/shape-opensheetwindow-method-visio%28Office.15%29.aspx)|
|[Paste](http://msdn.microsoft.com/library/shape-paste-method-visio%28Office.15%29.aspx)|
|[PasteSpecial](http://msdn.microsoft.com/library/shape-pastespecial-method-visio%28Office.15%29.aspx)|
|[RemoveFromContainers](http://msdn.microsoft.com/library/shape-removefromcontainers-method-visio%28Office.15%29.aspx)|
|[ReplaceShape](http://msdn.microsoft.com/library/shape-replaceshape-method-visio%28Office.15%29.aspx)|
|[Resize](http://msdn.microsoft.com/library/shape-resize-method-visio%28Office.15%29.aspx)|
|[ReverseEnds](http://msdn.microsoft.com/library/shape-reverseends-method-visio%28Office.15%29.aspx)|
|[Rotate90](http://msdn.microsoft.com/library/shape-rotate90-method-visio%28Office.15%29.aspx)|
|[SendBackward](http://msdn.microsoft.com/library/shape-sendbackward-method-visio%28Office.15%29.aspx)|
|[SendToBack](http://msdn.microsoft.com/library/shape-sendtoback-method-visio%28Office.15%29.aspx)|
|[SetBegin](http://msdn.microsoft.com/library/shape-setbegin-method-visio%28Office.15%29.aspx)|
|[SetCenter](http://msdn.microsoft.com/library/shape-setcenter-method-visio%28Office.15%29.aspx)|
|[SetEnd](http://msdn.microsoft.com/library/shape-setend-method-visio%28Office.15%29.aspx)|
|[SetFormulas](http://msdn.microsoft.com/library/shape-setformulas-method-visio%28Office.15%29.aspx)|
|[SetQuickStyle](http://msdn.microsoft.com/library/shape-setquickstyle-method-visio%28Office.15%29.aspx)|
|[SetResults](http://msdn.microsoft.com/library/shape-setresults-method-visio%28Office.15%29.aspx)|
|[SwapEnds](http://msdn.microsoft.com/library/shape-swapends-method-visio%28Office.15%29.aspx)|
|[TransformXYFrom](http://msdn.microsoft.com/library/shape-transformxyfrom-method-visio%28Office.15%29.aspx)|
|[TransformXYTo](http://msdn.microsoft.com/library/shape-transformxyto-method-visio%28Office.15%29.aspx)|
|[Ungroup](http://msdn.microsoft.com/library/shape-ungroup-method-visio%28Office.15%29.aspx)|
|[UpdateAlignmentBox](http://msdn.microsoft.com/library/shape-updatealignmentbox-method-visio%28Office.15%29.aspx)|
|[VisualBoundingBox](http://msdn.microsoft.com/library/shape-visualboundingbox-method-visio%28Office.15%29.aspx)|
|[XYFromPage](http://msdn.microsoft.com/library/shape-xyfrompage-method-visio%28Office.15%29.aspx)|
|[XYToPage](http://msdn.microsoft.com/library/shape-xytopage-method-visio%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/shape-application-property-visio%28Office.15%29.aspx)|
|[AreaIU](http://msdn.microsoft.com/library/shape-areaiu-property-visio%28Office.15%29.aspx)|
|[CalloutsAssociated](http://msdn.microsoft.com/library/shape-calloutsassociated-property-visio%28Office.15%29.aspx)|
|[CalloutTarget](http://msdn.microsoft.com/library/shape-callouttarget-property-visio%28Office.15%29.aspx)|
|[CellExists](http://msdn.microsoft.com/library/shape-cellexists-property-visio%28Office.15%29.aspx)|
|[CellExistsU](http://msdn.microsoft.com/library/shape-cellexistsu-property-visio%28Office.15%29.aspx)|
|[Cells](http://msdn.microsoft.com/library/shape-cells-property-visio%28Office.15%29.aspx)|
|[CellsRowIndex](http://msdn.microsoft.com/library/shape-cellsrowindex-property-visio%28Office.15%29.aspx)|
|[CellsRowIndexU](http://msdn.microsoft.com/library/shape-cellsrowindexu-property-visio%28Office.15%29.aspx)|
|[CellsSRC](http://msdn.microsoft.com/library/shape-cellssrc-property-visio%28Office.15%29.aspx)|
|[CellsSRCExists](http://msdn.microsoft.com/library/shape-cellssrcexists-property-visio%28Office.15%29.aspx)|
|[CellsU](http://msdn.microsoft.com/library/shape-cellsu-property-visio%28Office.15%29.aspx)|
|[Characters](http://msdn.microsoft.com/library/shape-characters-property-visio%28Office.15%29.aspx)|
|[CharCount](http://msdn.microsoft.com/library/shape-charcount-property-visio%28Office.15%29.aspx)|
|[ClassID](http://msdn.microsoft.com/library/shape-classid-property-visio%28Office.15%29.aspx)|
|[Comments](http://msdn.microsoft.com/library/shape-comments-property-visio%28Office.15%29.aspx)|
|[Connects](http://msdn.microsoft.com/library/shape-connects-property-visio%28Office.15%29.aspx)|
|[ContainerProperties](http://msdn.microsoft.com/library/shape-containerproperties-property-visio%28Office.15%29.aspx)|
|[ContainingMaster](http://msdn.microsoft.com/library/shape-containingmaster-property-visio%28Office.15%29.aspx)|
|[ContainingMasterID](http://msdn.microsoft.com/library/shape-containingmasterid-property-visio%28Office.15%29.aspx)|
|[ContainingPage](http://msdn.microsoft.com/library/shape-containingpage-property-visio%28Office.15%29.aspx)|
|[ContainingPageID](http://msdn.microsoft.com/library/shape-containingpageid-property-visio%28Office.15%29.aspx)|
|[ContainingShape](http://msdn.microsoft.com/library/shape-containingshape-property-visio%28Office.15%29.aspx)|
|[Data1](http://msdn.microsoft.com/library/shape-data1-property-visio%28Office.15%29.aspx)|
|[Data2](http://msdn.microsoft.com/library/shape-data2-property-visio%28Office.15%29.aspx)|
|[Data3](http://msdn.microsoft.com/library/shape-data3-property-visio%28Office.15%29.aspx)|
|[DataGraphic](http://msdn.microsoft.com/library/shape-datagraphic-property-visio%28Office.15%29.aspx)|
|[DistanceFrom](http://msdn.microsoft.com/library/shape-distancefrom-property-visio%28Office.15%29.aspx)|
|[DistanceFromPoint](http://msdn.microsoft.com/library/shape-distancefrompoint-property-visio%28Office.15%29.aspx)|
|[Document](http://msdn.microsoft.com/library/shape-document-property-visio%28Office.15%29.aspx)|
|[EventList](http://msdn.microsoft.com/library/shape-eventlist-property-visio%28Office.15%29.aspx)|
|[FillStyle](http://msdn.microsoft.com/library/shape-fillstyle-property-visio%28Office.15%29.aspx)|
|[FillStyleKeepFmt](http://msdn.microsoft.com/library/shape-fillstylekeepfmt-property-visio%28Office.15%29.aspx)|
|[ForeignData](http://msdn.microsoft.com/library/shape-foreigndata-property-visio%28Office.15%29.aspx)|
|[ForeignType](http://msdn.microsoft.com/library/shape-foreigntype-property-visio%28Office.15%29.aspx)|
|[FromConnects](http://msdn.microsoft.com/library/shape-fromconnects-property-visio%28Office.15%29.aspx)|
|[GeometryCount](http://msdn.microsoft.com/library/shape-geometrycount-property-visio%28Office.15%29.aspx)|
|[Help](http://msdn.microsoft.com/library/shape-help-property-visio%28Office.15%29.aspx)|
|[Hyperlinks](http://msdn.microsoft.com/library/shape-hyperlinks-property-visio%28Office.15%29.aspx)|
|[ID](http://msdn.microsoft.com/library/shape-id-property-visio%28Office.15%29.aspx)|
|[Index](http://msdn.microsoft.com/library/shape-index-property-visio%28Office.15%29.aspx)|
|[IsCallout](http://msdn.microsoft.com/library/shape-iscallout-property-visio%28Office.15%29.aspx)|
|[IsDataGraphicCallout](http://msdn.microsoft.com/library/shape-isdatagraphiccallout-property-visio%28Office.15%29.aspx)|
|[IsOpenForTextEdit](http://msdn.microsoft.com/library/shape-isopenfortextedit-property-visio%28Office.15%29.aspx)|
|[Language](http://msdn.microsoft.com/library/shape-language-property-visio%28Office.15%29.aspx)|
|[Layer](http://msdn.microsoft.com/library/shape-layer-property-visio%28Office.15%29.aspx)|
|[LayerCount](http://msdn.microsoft.com/library/shape-layercount-property-visio%28Office.15%29.aspx)|
|[LengthIU](http://msdn.microsoft.com/library/shape-lengthiu-property-visio%28Office.15%29.aspx)|
|[LineStyle](http://msdn.microsoft.com/library/shape-linestyle-property-visio%28Office.15%29.aspx)|
|[LineStyleKeepFmt](http://msdn.microsoft.com/library/shape-linestylekeepfmt-property-visio%28Office.15%29.aspx)|
|[Master](http://msdn.microsoft.com/library/shape-master-property-visio%28Office.15%29.aspx)|
|[MasterShape](http://msdn.microsoft.com/library/shape-mastershape-property-visio%28Office.15%29.aspx)|
|[MemberOfContainers](http://msdn.microsoft.com/library/shape-memberofcontainers-property-visio%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/shape-name-property-visio%28Office.15%29.aspx)|
|[NameID](http://msdn.microsoft.com/library/shape-nameid-property-visio%28Office.15%29.aspx)|
|[NameU](http://msdn.microsoft.com/library/shape-nameu-property-visio%28Office.15%29.aspx)|
|[Object](http://msdn.microsoft.com/library/shape-object-property-visio%28Office.15%29.aspx)|
|[ObjectIsInherited](http://msdn.microsoft.com/library/shape-objectisinherited-property-visio%28Office.15%29.aspx)|
|[ObjectType](http://msdn.microsoft.com/library/shape-objecttype-property-visio%28Office.15%29.aspx)|
|[OneD](http://msdn.microsoft.com/library/shape-oned-property-visio%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/shape-parent-property-visio%28Office.15%29.aspx)|
|[Paths](http://msdn.microsoft.com/library/shape-paths-property-visio%28Office.15%29.aspx)|
|[PathsLocal](http://msdn.microsoft.com/library/shape-pathslocal-property-visio%28Office.15%29.aspx)|
|[PersistsEvents](http://msdn.microsoft.com/library/shape-persistsevents-property-visio%28Office.15%29.aspx)|
|[Picture](http://msdn.microsoft.com/library/shape-picture-property-visio%28Office.15%29.aspx)|
|[ProgID](http://msdn.microsoft.com/library/shape-progid-property-visio%28Office.15%29.aspx)|
|[RootShape](http://msdn.microsoft.com/library/shape-rootshape-property-visio%28Office.15%29.aspx)|
|[RowCount](http://msdn.microsoft.com/library/shape-rowcount-property-visio%28Office.15%29.aspx)|
|[RowExists](http://msdn.microsoft.com/library/shape-rowexists-property-visio%28Office.15%29.aspx)|
|[RowsCellCount](http://msdn.microsoft.com/library/shape-rowscellcount-property-visio%28Office.15%29.aspx)|
|[RowType](http://msdn.microsoft.com/library/shape-rowtype-property-visio%28Office.15%29.aspx)|
|[Section](http://msdn.microsoft.com/library/shape-section-property-visio%28Office.15%29.aspx)|
|[SectionExists](http://msdn.microsoft.com/library/shape-sectionexists-property-visio%28Office.15%29.aspx)|
|[Shapes](http://msdn.microsoft.com/library/shape-shapes-property-visio%28Office.15%29.aspx)|
|[SpatialNeighbors](http://msdn.microsoft.com/library/shape-spatialneighbors-property-visio%28Office.15%29.aspx)|
|[SpatialRelation](http://msdn.microsoft.com/library/shape-spatialrelation-property-visio%28Office.15%29.aspx)|
|[SpatialSearch](http://msdn.microsoft.com/library/shape-spatialsearch-property-visio%28Office.15%29.aspx)|
|[Stat](http://msdn.microsoft.com/library/shape-stat-property-visio%28Office.15%29.aspx)|
|[Style](http://msdn.microsoft.com/library/shape-style-property-visio%28Office.15%29.aspx)|
|[StyleKeepFmt](http://msdn.microsoft.com/library/shape-stylekeepfmt-property-visio%28Office.15%29.aspx)|
|[Text](http://msdn.microsoft.com/library/shape-text-property-visio%28Office.15%29.aspx)|
|[TextStyle](http://msdn.microsoft.com/library/shape-textstyle-property-visio%28Office.15%29.aspx)|
|[TextStyleKeepFmt](http://msdn.microsoft.com/library/shape-textstylekeepfmt-property-visio%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/shape-type-property-visio%28Office.15%29.aspx)|
|[UniqueID](http://msdn.microsoft.com/library/shape-uniqueid-property-visio%28Office.15%29.aspx)|

