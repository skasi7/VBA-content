---
title: Master Members (Visio)
ms.prod: VISIO
ms.assetid: 8edaac93-fe1e-6163-ce16-3fec1fbb0fd2
---


# Master Members (Visio)

Represents a master in a stencil.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[BeforeMasterDelete](master-beforemasterdelete-event-visio.md)|Occurs before a master is deleted from a document.|
|[BeforeSelectionDelete](master-beforeselectiondelete-event-visio.md)|Occurs before selected objects are deleted.|
|[BeforeShapeDelete](master-beforeshapedelete-event-visio.md)|Occurs before a shape is deleted.|
|[BeforeShapeTextEdit](master-beforeshapetextedit-event-visio.md)|Occurs before a shape is opened for text editing in the user interface.|
|[CellChanged](master-cellchanged-event-visio.md)|Occurs after the value changes in a cell in a document.|
|[ConnectionsAdded](master-connectionsadded-event-visio.md)|Occurs after connections have been established between shapes.|
|[ConnectionsDeleted](master-connectionsdeleted-event-visio.md)|Occurs after connections between shapes have been removed.|
|[ConvertToGroupCanceled](master-converttogroupcanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelConvertToGroup** event.|
|[FormulaChanged](master-formulachanged-event-visio.md)|Occurs after a formula changes in a cell in the object that receives the event.|
|[GroupCanceled](master-groupcanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelGroup** event.|
|[MasterChanged](master-masterchanged-event-visio.md)|Occurs after properties of a master are changed and propagated to its instances.|
|[MasterDeleteCanceled](master-masterdeletecanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelMasterDelete** event.|
|[QueryCancelConvertToGroup](master-querycancelconverttogroup-event-visio.md)|Occurs before the application converts a selection of shapes to a group in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelGroup](master-querycancelgroup-event-visio.md)|Occurs before the application groups a selection of shapes in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelMasterDelete](master-querycancelmasterdelete-event-visio.md)|Occurs before the application deletes a master in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelSelectionDelete](master-querycancelselectiondelete-event-visio.md)|Occurs before the application deletes a selection of shapes in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelUngroup](master-querycancelungroup-event-visio.md)|Occurs before the application ungroups a selection of shapes in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[SelectionAdded](master-selectionadded-event-visio.md)|Occurs after one or more shapes are added to a document.|
|[SelectionDeleteCanceled](master-selectiondeletecanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelSelectionDelete** event.|
|[ShapeAdded](master-shapeadded-event-visio.md)|Occurs after one or more shapes are added to a document.|
|[ShapeChanged](master-shapechanged-event-visio.md)|Occurs after a property of a shape that is not stored in a cell is changed in a document.|
|[ShapeDataGraphicChanged](master-shapedatagraphicchanged-event-visio.md)|Occurs after a data graphic is applied to or deleted from a shape.|
|[ShapeExitedTextEdit](master-shapeexitedtextedit-event-visio.md)|Occurs after a shape is no longer open for interactive text editing.|
|[ShapeParentChanged](master-shapeparentchanged-event-visio.md)|Occurs after shapes are grouped or a group is ungrouped.|
|[TextChanged](master-textchanged-event-visio.md)|Occurs after the text of a shape is changed in a document.|
|[UngroupCanceled](master-ungroupcanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelUngroup** event.|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[AddGuide](master-addguide-method-visio.md)|Adds a guide to a master.|
|[BoundingBox](master-boundingbox-method-visio.md)|Returns a rectangle that tightly encloses the shapes of a master.|
|[CenterDrawing](master-centerdrawing-method-visio.md)|Centers a page's, master's, or group's shapes with respect to the extent of the page, master, or group. .|
|[Close](master-close-method-visio.md)|Closes a master.|
|[CreateSelection](master-createselection-method-visio.md)|Creates various types of  **Selection** objects.|
|[CreateShortcut](master-createshortcut-method-visio.md)|Creates a shortcut for a master.|
|[DataGraphicDelete](master-datagraphicdelete-method-visio.md)|Deletes the  **Master** of type **visTypeDataGraphic** from the **Masters** collection of the document.|
|[Delete](master-delete-method-visio.md)|Deletes an object.|
|[DrawArcByThreePoints](master-drawarcbythreepoints-method-visio.md)|Creates a shape whose path consists of an arc defined by the three points passed as parameters.|
|[DrawBezier](master-drawbezier-method-visio.md)|Creates a shape whose path is defined by the supplied sequence of Bezier control points.|
|[DrawCircularArc](master-drawcirculararc-method-visio.md)|Creates a new shape whose path consists of a circular arc defined by its center, radius, and start and end angles.|
|[DrawLine](master-drawline-method-visio.md)|Adds a line to the  **Shapes** collection of a master.|
|[DrawNURBS](master-drawnurbs-method-visio.md)|Creates a new shape whose path consists of a single NURBS (nonuniform rational B-spline) segment.|
|[DrawOval](master-drawoval-method-visio.md)|Adds an oval (ellipse) to the  **Shapes** collection of a master.|
|[DrawPolyline](master-drawpolyline-method-visio.md)|Creates a shape whose path is a polyline along a given set of points.|
|[DrawQuarterArc](master-drawquarterarc-method-visio.md)|Creates a shape whose path consists of an elliptical arc defined by the two points and the flag passed in as arguments.|
|[DrawRectangle](master-drawrectangle-method-visio.md)|Adds a rectangle to the  **Shapes** collection of a page, master, or group.|
|[DrawSpline](master-drawspline-method-visio.md)|Creates a new shape whose path follows a given sequence of points.|
|[Drop](master-drop-method-visio.md)|Creates one or more new  **Shape** objects by dropping an object onto a receiving object such as a master, drawing page, shape, or group.|
|[DropMany](master-dropmany-method-visio.md)|Creates one or more new  **Shape** objects in a master. It returns an array of the IDs of the **Shape** objects it produces.|
|[DropManyU](master-dropmanyu-method-visio.md)|Creates one or more new  **Shape** objects on a page, in a master, or in a group. It returns an array of the IDs of the **Shape** objects it produces.|
|[Export](master-export-method-visio.md)|Exports an object from Microsoft Visio to a file format such as .bmp, .dib, .dwg, .dxf, .emf, .emz, .gif, .htm, .jpg, .png, .svg, .svgz, .tif, or .wmf.|
|[ExportIcon](master-exporticon-method-visio.md)|Exports the icon for a  **Master** object to a named file or the Clipboard.|
|[GetFormulas](master-getformulas-method-visio.md)|Returns the formulas of many cells.|
|[GetFormulasU](master-getformulasu-method-visio.md)|Returns the formulas of many cells.|
|[GetResults](master-getresults-method-visio.md)|Gets the results or formulas of many cells.|
|[Import](master-import-method-visio.md)|Imports a file into the current document.|
|[ImportIcon](master-importicon-method-visio.md)|Imports the icon for a  **Master** object from a named file.|
|[InsertFromFile](master-insertfromfile-method-visio.md)|Adds a linked or embedded object to a page, master, or group.|
|[InsertObject](master-insertobject-method-visio.md)|Adds a new embedded object or ActiveX control to a page, master, or group.|
|[Layout](master-layout-method-visio.md)|Lays out the shapes and/or reroutes the connectors for the page, master, group, or selection.|
|[Open](master-open-method-visio.md)|Opens an existing master so that it can be edited.|
|[OpenDrawWindow](master-opendrawwindow-method-visio.md)|Opens a new drawing window that displays a master.|
|[OpenIconWindow](master-openiconwindow-method-visio.md)|Opens an icon window that shows a master's icon.|
|[Paste](master-paste-method-visio.md)|Pastes the contents of the Clipboard into an object.|
|[PasteSpecial](master-pastespecial-method-visio.md)|Inserts the contents of the Clipboard, allowing you to control the format of the pasted information and (optionally) establish a link to the source file (for example, a Microsoft Word document).|
|[PasteToLocation](master-pastetolocation-method-visio.md)|Pastes a shape to the specified location.|
|[ResizeToFitContents](master-resizetofitcontents-method-visio.md)|Resizes the page, or the master's page, to fit tightly around the shapes or master that are on it.|
|[SetFormulas](master-setformulas-method-visio.md)|Sets the formulas of one or more cells.|
|[SetResults](master-setresults-method-visio.md)|Sets the results or formulas of one or more cells.|
|[VisualBoundingBox](master-visualboundingbox-method-visio.md)|Returns the bounding rectangle of the virtual container that has all the shapes of the given master.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[AlignName](master-alignname-property-visio.md)|Gets or sets the position of a master name in a stencil window. Read/write.|
|[Application](master-application-property-visio.md)|Returns the instance of Microsoft Visio that is associated with an object. Read-only.|
|[BaseID](master-baseid-property-visio.md)|Returns a base ID for a master. Read-only.|
|[Connects](master-connects-property-visio.md)|Returns a  **Connects** collection for a shape, page, or master. Read-only.|
|[DataGraphicHidden](master-datagraphichidden-property-visio.md)|Hides or displays a data graphic in the  **Data Graphics** task pane in the Microsoft Visio user interface. Read/write.|
|[DataGraphicHidesText](master-datagraphichidestext-property-visio.md)|Displays or hides the text of a shape or of the primary shape in a selection when a data graphic is applied to the shape or to the selection. Read/write.|
|[DataGraphicHorizontalPosition](master-datagraphichorizontalposition-property-visio.md)|Gets or sets the default horizontal callout position for members of the  **GraphicItems** collection of the **Master** object of type **visTypeDataGraphic** . Read/write.|
|[DataGraphicShowBorder](master-datagraphicshowborder-property-visio.md)|Gets or sets whether a border is displayed around the graphic items contained in the data graphic that are in default positions. Read/write.|
|[DataGraphicVerticalPosition](master-datagraphicverticalposition-property-visio.md)|Gets or sets the default vertical callout position for members of the  **GraphicItems** collection of the **Master** object of type **visTypeDataGraphic** . Read/write.|
|[Document](master-document-property-visio.md)|Gets the  **Document** object that is associated with an object. Read-only.|
|[EditCopy](master-editcopy-property-visio.md)|Returns a master that is open for editing and originally copied from this master. Read-only.|
|[EventList](master-eventlist-property-visio.md)|Returns the  **EventList** collection of an object or the **EventList** collection that contains an **Event** object. Read-only.|
|[GraphicItems](master-graphicitems-property-visio.md)|Returns the  **GraphicItems** collection that the master contains. Read-only.|
|[Hidden](master-hidden-property-visio.md)|Hides or shows a master on a stencil or a style in the user interface. Read/write.|
|[Icon](master-icon-property-visio.md)|Returns the icon contained in a master. Read/write.|
|[IconSize](master-iconsize-property-visio.md)|Gets or sets the size of a master icon. Read/write.|
|[IconUpdate](master-iconupdate-property-visio.md)|Determines whether a master icon is updated manually or automatically. Read/write.|
|[ID](master-id-property-visio.md)|Gets the ID of an object. Read-only.|
|[Index](master-index-property-visio.md)|Gets the ordinal position of a  **Master** object in the **Masters** collection. Read-only.|
|[IndexInStencil](master-indexinstencil-property-visio.md)|Gets or sets the index of a master or master shortcut object within its stencil. Read/write.|
|[IsChanged](master-ischanged-property-visio.md)|Indicates whether a master has changed since it was opened. Read-only.|
|[Layers](master-layers-property-visio.md)|Returns the  **Layers** collection of an object. Read-only.|
|[MatchByName](master-matchbyname-property-visio.md)|Determines how the application decides if a document master is already present when an instance of a master is dropped on the drawing page. It allows changes made to a document master to apply to new instances of the master, even if the instances are dragged from a stand-alone stencil file. Read/write.|
|[Name](master-name-property-visio.md)|Specifies the name of an object. Read-only.|
|[NameU](master-nameu-property-visio.md)|Specifies the universal name of a  **Master** object. Read/write.|
|[NewBaseID](master-newbaseid-property-visio.md)|Generates a new base ID for a master. Read-only.|
|[ObjectType](master-objecttype-property-visio.md)|Returns an object's type. Read-only.|
|[OLEObjects](master-oleobjects-property-visio.md)|Returns the  **OLEObjects** collection of a master. Read-only.|
|[OneD](master-oned-property-visio.md)|Determines whether an object behaves as a one-dimensional (1-D) object. Read-only.|
|[Original](master-original-property-visio.md)|Returns the original master that produced this open master. Read-only.|
|[PageSheet](master-pagesheet-property-visio.md)|Returns the page sheet (an object that represents the ShapeSheet spreadsheet) of a master. Read-only.|
|[PatternFlags](master-patternflags-property-visio.md)|Determines whether a master behaves as a custom pattern. Read/write.|
|[PersistsEvents](master-persistsevents-property-visio.md)|Indicates whether an object is capable of containing persistent events in its  **EventList** collection. Read-only.|
|[Picture](master-picture-property-visio.md)|Returns a picture that represents an enhanced metafile (EMF) contained in a master, shape, selection, or page. Read-only.|
|[Prompt](master-prompt-property-visio.md)|Gets or sets the prompt string for a master or master shortcut. Read/write.|
|[Shapes](master-shapes-property-visio.md)|Returns the  **Shapes** collection for a page, master, or group. Read-only.|
|[SpatialSearch](master-spatialsearch-property-visio.md)|Returns a  **Selection** object whose shapes meet certain criteria in relation to a point that is expressed in the coordinate space of a page, master, or group. Read-only.|
|[Stat](master-stat-property-visio.md)|Returns status information for an object. Read-only.|
|[Type](master-type-property-visio.md)|Returns the type of the  **Master** object. Read-only.|
|[UniqueID](master-uniqueid-property-visio.md)|Returns the unique ID of a master. Read-only.|

