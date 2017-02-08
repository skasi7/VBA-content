---
title: Shape Members (Visio)
ms.prod: VISIO
ms.assetid: 6aee2782-bec8-b9fd-bfa6-4c30c1dec8eb
---


# Shape Members (Visio)
Represents anything you can select in a drawing window: a basic shape, a group, a guide, or an object from another application embedded or linked in Microsoft Visio.

Represents anything you can select in a drawing window: a basic shape, a group, a guide, or an object from another application embedded or linked in Microsoft Visio.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[BeforeSelectionDelete](shape-beforeselectiondelete-event-visio.md)|Occurs before selected objects are deleted.|
|[BeforeShapeDelete](shape-beforeshapedelete-event-visio.md)|Occurs before a shape is deleted.|
|[BeforeShapeTextEdit](shape-beforeshapetextedit-event-visio.md)|Occurs before a shape is opened for text editing in the user interface.|
|[CellChanged](shape-cellchanged-event-visio.md)|Occurs after the value changes in a cell in a document.|
|[ConvertToGroupCanceled](shape-converttogroupcanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelConvertToGroup** event.|
|[FormulaChanged](shape-formulachanged-event-visio.md)|Occurs after a formula changes in a cell in the object that receives the event.|
|[GroupCanceled](shape-groupcanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelGroup** event.|
|[QueryCancelConvertToGroup](shape-querycancelconverttogroup-event-visio.md)|Occurs before the application converts a selection of shapes to a group in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelGroup](shape-querycancelgroup-event-visio.md)|Occurs before the application groups a selection of shapes in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelSelectionDelete](shape-querycancelselectiondelete-event-visio.md)|Occurs before the application deletes a selection of shapes in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelUngroup](shape-querycancelungroup-event-visio.md)|Occurs before the application ungroups a selection of shapes in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[SelectionAdded](shape-selectionadded-event-visio.md)|Occurs after one or more shapes are added to a document.|
|[SelectionDeleteCanceled](shape-selectiondeletecanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelSelectionDelete** event.|
|[ShapeAdded](shape-shapeadded-event-visio.md)|Occurs after one or more shapes are added to a document.|
|[ShapeChanged](shape-shapechanged-event-visio.md)|Occurs after a property of a shape that is not stored in a cell is changed in a document.|
|[ShapeDataGraphicChanged](shape-shapedatagraphicchanged-event-visio.md)|Occurs after a data graphic is applied to or deleted from a shape.|
|[ShapeExitedTextEdit](shape-shapeexitedtextedit-event-visio.md)|Occurs after a shape is no longer open for interactive text editing.|
|[ShapeLinkAdded](shape-shapelinkadded-event-visio.md)|Occurs after a shape is linked to a data row.|
|[ShapeLinkDeleted](shape-shapelinkdeleted-event-visio.md)|Occurs after the link between a shape and a data row is deleted.|
|[ShapeParentChanged](shape-shapeparentchanged-event-visio.md)|Occurs after shapes are grouped or a group is ungrouped.|
|[TextChanged](shape-textchanged-event-visio.md)|Occurs after the text of a shape is changed in a document.|
|[UngroupCanceled](shape-ungroupcanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelUngroup** event.|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[AddGuide](shape-addguide-method-visio.md)|Adds a guide to a group shape.|
|[AddHyperlink](shape-addhyperlink-method-visio.md)|Adds a  **Hyperlink** object to a Microsoft Visio shape.|
|[AddNamedRow](shape-addnamedrow-method-visio.md)|Adds a row that has the specified name to the specified ShapeSheet section.|
|[AddRow](shape-addrow-method-visio.md)|Adds a row to a ShapeSheet section at a specified position.|
|[AddRows](shape-addrows-method-visio.md)|Adds the specified number of rows to a ShapeSheet section at a specified position.|
|[AddSection](shape-addsection-method-visio.md)|Adds a new section to a ShapeSheet spreadsheet.|
|[AddToContainers](shape-addtocontainers-method-visio.md)|Adds the shape to all underlying containers that allow it as a member.|
|[AutoConnect](shape-autoconnect-method-visio.md)|Automatically draws a connection in the specified direction between the shape and another shape on the drawing page.|
|[BoundingBox](shape-boundingbox-method-visio.md)|Returns a rectangle that tightly encloses a shape.|
|[BreakLinkToData](shape-breaklinktodata-method-visio.md)|Breaks the link between the shape and the data row to which it is linked in the specified data recordset.|
|[BringForward](shape-bringforward-method-visio.md)|Brings the shape or selected shapes forward one position in the z-order.|
|[BringToFront](shape-bringtofront-method-visio.md)|Brings the shape or selected shapes to the front of the z-order.|
|[CenterDrawing](shape-centerdrawing-method-visio.md)|Centers a page's, master's, or group's shapes with respect to the extent of the page, master, or group. .|
|[ChangePicture](shape-changepicture-method-visio.md)|Replaces the specified shape?s current picture with a new picture.|
|[ConnectedShapes](shape-connectedshapes-method-visio.md)|Returns an array that contains the identifiers (IDs) of the shapes that are connected to the shape.|
|[ConvertToGroup](shape-converttogroup-method-visio.md)|Converts a selection or an object from another application (a linked or embedded object) to a group.|
|[Copy](shape-copy-method-visio.md)|Copies a shape to the Clipboard.|
|[CreateSelection](shape-createselection-method-visio.md)|Creates various types of  **Selection** objects.|
|[CreateSubProcess](shape-createsubprocess-method-visio.md)|Creates and returns a new sub-process page that is linked to the shape.|
|[Cut](shape-cut-method-visio.md)|Deletes an object or selection and places it on the Clipboard.|
|[Delete](shape-delete-method-visio.md)|Deletes an object or selection.|
|[DeleteEx](shape-deleteex-method-visio.md)|Deletes the additional shapes that are associated with the shape, such as connectors and unselected container members, when the shape is deleted.|
|[DeleteRow](shape-deleterow-method-visio.md)|Deletes a row from a section in a ShapeSheet spreadsheet.|
|[DeleteSection](shape-deletesection-method-visio.md)|Deletes a ShapeSheet section.|
|[Disconnect](shape-disconnect-method-visio.md)|Unglues the specified connector end points and offsets them the specified amount from the shapes to which they were joined.|
|[DrawArcByThreePoints](shape-drawarcbythreepoints-method-visio.md)|Creates a shape whose path consists of an arc defined by the three points passed as parameters.|
|[DrawBezier](shape-drawbezier-method-visio.md)|Creates a shape whose path is defined by the supplied sequence of Bezier control points.|
|[DrawCircularArc](shape-drawcirculararc-method-visio.md)|Creates a new shape whose path consists of a circular arc defined by its center, radius, and start and end angles.|
|[DrawLine](shape-drawline-method-visio.md)|Adds a line to the  **Shapes** collection of a group shape.|
|[DrawNURBS](shape-drawnurbs-method-visio.md)|Creates a new shape whose path consists of a single NURBS (nonuniform rational B-spline) segment.|
|[DrawOval](shape-drawoval-method-visio.md)|Adds an oval (ellipse) to the  **Shapes** collection of a group shape.|
|[DrawPolyline](shape-drawpolyline-method-visio.md)|Creates a shape whose path is a polyline along a given set of points.|
|[DrawQuarterArc](shape-drawquarterarc-method-visio.md)|Creates a new shape whose path consists of an elliptical arc defined by the two points and the flag passed in as arguments.|
|[DrawRectangle](shape-drawrectangle-method-visio.md)|Adds a rectangle to the  **Shapes** collection of a page, master, or group.|
|[DrawSpline](shape-drawspline-method-visio.md)|Creates a new shape whose path follows a given sequence of points.|
|[Drop](shape-drop-method-visio.md)|Creates one or more new  **Shape** objects by dropping an object onto a receiving object such as a master, drawing page, shape, or group.|
|[DropMany](shape-dropmany-method-visio.md)|Creates one or more new  **Shape** objects in a group. It returns an array of the IDs of the **Shape** objects it produces.|
|[DropManyU](shape-dropmanyu-method-visio.md)|Creates one or more new  **Shape** objects on a page, in a master, or in a group. It returns an array of the IDs of the **Shape** objects it produces.|
|[Duplicate](shape-duplicate-method-visio.md)|Duplicates an object.|
|[Export](shape-export-method-visio.md)|Exports an object from Microsoft Visio to a file format such as .bmp, .dib, .dwg, .dxf, .emf, .emz, .gif, .htm, .jpg, .png, .svg, .svgz, .tif, or .wmf.|
|[FitCurve](shape-fitcurve-method-visio.md)|Reduces the number of geometry segments in a shape or shapes by replacing them with similar spline, arc, and line segments that approximate the paths of the initial segments. Typically, this reduces the number of segments in the shape.|
|[FlipHorizontal](shape-fliphorizontal-method-visio.md)|Flips an object horizontally.|
|[FlipVertical](shape-flipvertical-method-visio.md)|Flips an object vertically.|
|[GetCustomPropertiesLinkedToData](shape-getcustompropertieslinkedtodata-method-visio.md)|Gets the IDs of the shape-data-item (custom property) rows in the Shape Data section of the shape's ShapeSheet spreadsheet linked to the specified data recordset.|
|[GetCustomPropertyLinkedColumn](shape-getcustompropertylinkedcolumn-method-visio.md)|Gets the name of the data column linked to the shape data (custom property) row in the shape's ShapeSheet spreadsheet specified by the custom property index.|
|[GetFormulas](shape-getformulas-method-visio.md)|Returns the formulas of many cells.|
|[GetFormulasU](shape-getformulasu-method-visio.md)|Returns the formulas of many cells.|
|[GetLinkedDataRecordsetIDs](shape-getlinkeddatarecordsetids-method-visio.md)|Gets the IDs of all the data recordsets that contain data rows linked to the shape.|
|[GetLinkedDataRow](shape-getlinkeddatarow-method-visio.md)|Gets the ID of the data row in the specified data recordset linked to the shape.|
|[GetResults](shape-getresults-method-visio.md)|Gets the results or formulas of many cells.|
|[GluedShapes](shape-gluedshapes-method-visio.md)|Returns an array that contains the identifiers of the shapes that are glued to a shape.|
|[Group](shape-group-method-visio.md)|Groups the objects that are selected in a selection, or it converts a shape into a group.|
|[HasCategory](shape-hascategory-method-visio.md)|Returns  **True** if the specified category is in the shape categories list.|
|[HitTest](shape-hittest-method-visio.md)|Determines if a given  _x,y_ position hits outside, inside, or on the boundary of a shape.|
|[Import](shape-import-method-visio.md)|Imports a file into the current document.|
|[InsertFromFile](shape-insertfromfile-method-visio.md)|Adds a linked or embedded object to a page, master, or group.|
|[InsertObject](shape-insertobject-method-visio.md)|Adds a new embedded object or ActiveX control to a page, master, or group.|
|[IsCustomPropertyLinked](shape-iscustompropertylinked-method-visio.md)|Returns whether the shape data (custom property) row in the Shape Data section of the shape's ShapeSheet spreadsheet is linked to a data row in the specified data recordset.|
|[Layout](shape-layout-method-visio.md)|Lays out the shapes or reroutes the connectors (or both) for the page, master, group, or selection.|
|[LinkToData](shape-linktodata-method-visio.md)|Links a shape to a data row in a data recordset.|
|[MoveToSubprocess](shape-movetosubprocess-method-visio.md)|Moves the shape to the specified page and drops a replacement shape on the source page, then links it to the target page. Returns the selection of moved shapes on the target page.|
|[Offset](shape-offset-method-visio.md)|Offsets a shape a specified amount.|
|[OpenDrawWindow](shape-opendrawwindow-method-visio.md)|Opens a new drawing window that displays a group.|
|[OpenSheetWindow](shape-opensheetwindow-method-visio.md)|Opens a ShapeSheet window for a  **Shape** object.|
|[Paste](shape-paste-method-visio.md)|Pastes the contents of the Clipboard into an object.|
|[PasteSpecial](shape-pastespecial-method-visio.md)|Inserts the contents of the Clipboard, allowing you to control the format of the pasted information and (optionally) establish a link to the source file (for example, a Microsoft Word document).|
|[RemoveFromContainers](shape-removefromcontainers-method-visio.md)|Removes the shape from all lists and containers of which it is a member.|
|[ReplaceShape](shape-replaceshape-method-visio.md)|Replaces the specified shape with an instance of the master passed as the first parameter, and returns the new shape.|
|[Resize](shape-resize-method-visio.md)|Resizes the shape by moving shape handles as specified.|
|[ReverseEnds](shape-reverseends-method-visio.md)|Reverses an object by flipping it both horizontally and vertically.|
|[Rotate90](shape-rotate90-method-visio.md)|Rotates an object 90 degrees counterclockwise.|
|[SendBackward](shape-sendbackward-method-visio.md)|Moves a shape or selected shapes back one position in the z-order.|
|[SendToBack](shape-sendtoback-method-visio.md)|Moves the shape or selected shapes to the back of the z-order.|
|[SetBegin](shape-setbegin-method-visio.md)|Moves the begin point of a one-dimensional (1-D) shape to the coordinates represented by  _xPos_ and _yPos_.|
|[SetCenter](shape-setcenter-method-visio.md)|Moves a shape so that its pin is positioned at the coordinates represented by  _xPos_ and _yPos_. .|
|[SetEnd](shape-setend-method-visio.md)|Moves the endpoint of a one-dimensional (1-D) shape to the coordinates represented by  _xPos_ and _yPos_.|
|[SetFormulas](shape-setformulas-method-visio.md)|Sets the formulas of one or more cells.|
|[SetQuickStyle](shape-setquickstyle-method-visio.md)|Sets the quick style of the specified shape.|
|[SetResults](shape-setresults-method-visio.md)|Sets the results or formulas of one or more cells.|
|[SwapEnds](shape-swapends-method-visio.md)|Swaps the begin and endpoints of a one-dimensional (1-D) shape.|
|[TransformXYFrom](shape-transformxyfrom-method-visio.md)|Transforms a point expressed in the local coordinate system of one  **Shape** object from an equivalent point expressed in the local coordinate system of another **Shape** object.|
|[TransformXYTo](shape-transformxyto-method-visio.md)|Transforms a point expressed in the local coordinate system of one  **Shape** object to an equivalent point expressed in the local coordinate system of another **Shape** object.|
|[Ungroup](shape-ungroup-method-visio.md)|Ungroups a group.|
|[UpdateAlignmentBox](shape-updatealignmentbox-method-visio.md)|Updates the alignment box for a shape.|
|[VisualBoundingBox](shape-visualboundingbox-method-visio.md)|Returns the bounding rectangle of the given shape.|
|[XYFromPage](shape-xyfrompage-method-visio.md)|Transforms a point expressed in the local coordinate system of its  **Page** or **Master** object to an equivalent point expressed in the local coordinate system of the **Shape** object.|
|[XYToPage](shape-xytopage-method-visio.md)|Transforms a point expressed in the local coordinate system of a  **Shape** object to an equivalent point expressed in the local coordinate system of its **Page** or **Master** object.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](shape-application-property-visio.md)|Returns the instance of Microsoft Visio that is associated with an object. Read-only.|
|[AreaIU](shape-areaiu-property-visio.md)|Returns the area of a  **Shape** object in internal units (square inches). Read-only.|
|[CalloutsAssociated](shape-calloutsassociated-property-visio.md)|Returns an array of  **Long** values that represent the collection of identifiers for all of the callout shapes that are associated with the target shape by a callout relationship. Read-only.|
|[CalloutTarget](shape-callouttarget-property-visio.md)|Gets or sets the target shape that is associated with the callout shape by a callout relationship. Read/write.|
|[CellExists](shape-cellexists-property-visio.md)|Determines whether a particular ShapeSheet cell exists in the scope of the search. Read-only.|
|[CellExistsU](shape-cellexistsu-property-visio.md)|Determines whether a particular ShapeSheet cell exists in the scope of the search. Read-only.|
|[Cells](shape-cells-property-visio.md)|Returns a  **Cell** object that represents a ShapeSheet cell. Read-only.|
|[CellsRowIndex](shape-cellsrowindex-property-visio.md)|Returns the index of a row to which a cell belongs. Read-only.|
|[CellsRowIndexU](shape-cellsrowindexu-property-visio.md)|Returns the index of a row to which a cell belongs. Read-only.|
|[CellsSRC](shape-cellssrc-property-visio.md)|Returns a  **Cell** object that represents a ShapeSheet cell identified by section, row, and column indices. Read-only.|
|[CellsSRCExists](shape-cellssrcexists-property-visio.md)|Determines whether a ShapeSheet cell exists in the scope of a search. Read-only.|
|[CellsU](shape-cellsu-property-visio.md)|Returns a  **Cell** object that represents a ShapeSheet cell. Read-only.|
|[Characters](shape-characters-property-visio.md)|Returns a  **Characters** object that represents the text of a shape. Read-only.|
|[CharCount](shape-charcount-property-visio.md)|Returns the number of characters in an object. Read-only.|
|[ClassID](shape-classid-property-visio.md)|Returns the class ID string of a shape that represents an ActiveX control or an embedded or linked OLE object. Read-only.|
|[Comments](shape-comments-property-visio.md)|Returns a [Comments](comments-object-visio.md) object that represents the collection of all the reviewer comments on the shape. Read-only.|
|[Connects](shape-connects-property-visio.md)|Returns a  **Connects** collection for a shape, page, or master. Read-only.|
|[ContainerProperties](shape-containerproperties-property-visio.md)|Returns the  **[ContainerProperties](containerproperties-object-visio.md)** object associated with the shape. Read-only.|
|[ContainingMaster](shape-containingmaster-property-visio.md)|Returns the  **Master** object that contains an object. Read-only.|
|[ContainingMasterID](shape-containingmasterid-property-visio.md)|Returns the ID of the  **Master** object that contains an object. Read-only.|
|[ContainingPage](shape-containingpage-property-visio.md)|Returns the page that contains an object.|
|[ContainingPageID](shape-containingpageid-property-visio.md)|Returns the ID of the page that contains an object. Read-only.|
|[ContainingShape](shape-containingshape-property-visio.md)|Returns the  **Shape** object that contains an object or collection. Read-only.|
|[Data1](shape-data1-property-visio.md)|Gets or sets the value of the  **Data1** field for a **Shape** object. Read/write.|
|[Data2](shape-data2-property-visio.md)|Gets or sets the value of the  **Data2** field for a **Shape** object. Read/write.|
|[Data3](shape-data3-property-visio.md)|Gets or sets the value of the  **Data3** field for a **Shape** object. Read/write.|
|[DataGraphic](shape-datagraphic-property-visio.md)|Gets or sets the data graphic master ( **Master** of type **visTypeDataGraphic** ) that is associated with the shape. Read/write.|
|[DistanceFrom](shape-distancefrom-property-visio.md)|Returns the distance from one shape to another, measured between the closest points on the two shapes. Both shapes must be on the same page or in the same master. Read-only.|
|[DistanceFromPoint](shape-distancefrompoint-property-visio.md)|Returns the distance from a shape to a point. Read-only.|
|[Document](shape-document-property-visio.md)|Gets the  **Document** object that is associated with an object. Read-only.|
|[EventList](shape-eventlist-property-visio.md)|Returns the  **EventList** collection of an object or the **EventList** collection that contains an **Event** object. Read-only.|
|[FillStyle](shape-fillstyle-property-visio.md)|Returns or sets the fill style for an shape. Read/write.|
|[FillStyleKeepFmt](shape-fillstylekeepfmt-property-visio.md)|Applies a fill style to an object while preserving local formatting. Read/write.|
|[ForeignData](shape-foreigndata-property-visio.md)|Returns metafile, bitmap, or OLE data for a shape that represents a foreign object. Read-only.|
|[ForeignType](shape-foreigntype-property-visio.md)|Returns the subtype of a  **Shape** object that represents a foreign object. Read-only.|
|[FromConnects](shape-fromconnects-property-visio.md)|Returns a  **Connects** collection of the shapes connected to a shape. Read-only.|
|[GeometryCount](shape-geometrycount-property-visio.md)|Returns the number of Geometry sections for a shape. Read-only.|
|[Help](shape-help-property-visio.md)|Gets or sets the Help string for a shape. Read/write.|
|[Hyperlinks](shape-hyperlinks-property-visio.md)|Returns the  **Hyperlinks** collection for a **Shape** object. Read-only.|
|[ID](shape-id-property-visio.md)|Gets the ID of an object. Read-only.|
|[Index](shape-index-property-visio.md)|Gets the ordinal position of a  **Shape** object in the **Shapes** collection. Read-only.|
|[IsCallout](shape-iscallout-property-visio.md)|Indicates whether the shape is a callout shape. Read-only.|
|[IsDataGraphicCallout](shape-isdatagraphiccallout-property-visio.md)|Specifes whether a shape is a data graphic callout. Read-only.|
|[IsOpenForTextEdit](shape-isopenfortextedit-property-visio.md)|Indicates whether a shape is currently open for interactive text editing. Read-only.|
|[Language](shape-language-property-visio.md)|Represents the language ID of the version of the Microsoft Visio instance represented by the parent object. Read/write.|
|[Layer](shape-layer-property-visio.md)|Returns the layer to which a shape is assigned. Read-only.|
|[LayerCount](shape-layercount-property-visio.md)|Returns the number of layers to which a shape is assigned. Read-only.|
|[LengthIU](shape-lengthiu-property-visio.md)|Returns the length (perimeter) of the shape in internal units. Read-only.|
|[LineStyle](shape-linestyle-property-visio.md)|Specifies the line style for an object. Read/write.|
|[LineStyleKeepFmt](shape-linestylekeepfmt-property-visio.md)|Applies a line style to an object while preserving local formatting. Read/write.|
|[Master](shape-master-property-visio.md)|Returns the master from which the  **Shape** object was created. Read-only.|
|[MasterShape](shape-mastershape-property-visio.md)|If this shape is part of a master instance, returns the shape in the master that this shape inherits from. Read-only.|
|[MemberOfContainers](shape-memberofcontainers-property-visio.md)|Returns an array of  **Long** values that represent the identifiers of the container shapes that include the shape as a member. Read-only.|
|[Name](shape-name-property-visio.md)|Specifies the name of an object. Read-only.|
|[NameID](shape-nameid-property-visio.md)|Returns a unique name for a shape. Read-only.|
|[NameU](shape-nameu-property-visio.md)|Specifies the universal name of a  **Shape** object. Read/write.|
|[Object](shape-object-property-visio.md)|Returns an  **IDispatch** interface on the ActiveX control or embedded or linked OLE 2.0 object represented by a **Shape** object or an **OLEObject** object. Read-only.|
|[ObjectIsInherited](shape-objectisinherited-property-visio.md)|Indicates if a shape represents an ActiveX or OLE object that is inherited from the shape's master. Read-only.|
|[ObjectType](shape-objecttype-property-visio.md)|Returns an object's type. Read-only.|
|[OneD](shape-oned-property-visio.md)|Determines whether an object behaves as a one-dimensional (1-D) object. Read-only.|
|[Parent](shape-parent-property-visio.md)|Determines the parent of a  **Shape** object. Read/write.|
|[Paths](shape-paths-property-visio.md)|Returns a  **Paths** collection that reports the coordinates of a shape's paths in the coordinate system of the shape's parent. Read-only.|
|[PathsLocal](shape-pathslocal-property-visio.md)|Returns a  **Paths** collection that reports the coordinates of a shape's paths in the shape's local coordinate system. Read-only.|
|[PersistsEvents](shape-persistsevents-property-visio.md)|Indicates whether an object is capable of containing persistent events in its  **EventList** collection. Read-only.|
|[Picture](shape-picture-property-visio.md)|Returns a picture that represents an enhanced metafile (EMF) contained in a master, shape, selection, or page. Read-only.|
|[ProgID](shape-progid-property-visio.md)|Returns the programmatic identifier of a shape that represents an ActiveX control, an embedded object, or linked object. Read-only.|
|[RootShape](shape-rootshape-property-visio.md)|Returns the top-level shape of an instance if this shape is part of a master instance. Read-only.|
|[RowCount](shape-rowcount-property-visio.md)|Returns the number of rows in a ShapeSheet section. Read-only.|
|[RowExists](shape-rowexists-property-visio.md)|Determines whether a ShapeSheet row exists. Read-only.|
|[RowsCellCount](shape-rowscellcount-property-visio.md)|Returns the number of cells in a row of a ShapeSheet section. Read-only.|
|[RowType](shape-rowtype-property-visio.md)|Gets or sets the type of a row in a Geometry, Connection Points, Controls, or Tabs ShapeSheet section. Read/write.|
|[Section](shape-section-property-visio.md)|Returns the requested  **Section** object belonging to a shape. Read-only.|
|[SectionExists](shape-sectionexists-property-visio.md)|Determines whether a ShapeSheet section exists for a particular shape. Read-only.|
|[Shapes](shape-shapes-property-visio.md)|Returns the  **Shapes** collection for a page, master, or group. Read-only.|
|[SpatialNeighbors](shape-spatialneighbors-property-visio.md)|Returns a  **Selection** object that represents the shapes that meet certain criteria in relation to a specified shape. Read-only.|
|[SpatialRelation](shape-spatialrelation-property-visio.md)|Returns an integer that represents the spatial relationship of one shape to another shape. Both shapes must be on the same page or in the same master. Read-only.|
|[SpatialSearch](shape-spatialsearch-property-visio.md)|Returns a  **Selection** object whose shapes meet certain criteria in relation to a point that is expressed in the coordinate space of a page, master, or group. Read-only.|
|[Stat](shape-stat-property-visio.md)|Returns status information for an object. Read-only.|
|[Style](shape-style-property-visio.md)|Gets or sets the style for a  **Shape** object. Read/write.|
|[StyleKeepFmt](shape-stylekeepfmt-property-visio.md)|Applies a style to an object while preserving local formatting. Read/write.|
|[Text](shape-text-property-visio.md)|Returns all of the shape's text. Read/write.|
|[TextStyle](shape-textstyle-property-visio.md)|Gets or sets the text style for an object. Read/write.|
|[TextStyleKeepFmt](shape-textstylekeepfmt-property-visio.md)|Applies a text style to an object while preserving local formatting. Read/write.|
|[Type](shape-type-property-visio.md)|Returns the type of the object. Read-only.|
|[UniqueID](shape-uniqueid-property-visio.md)|Gets, deletes, or makes the GUID that uniquely identifies the shape within the scope of the application. Read-only.|

