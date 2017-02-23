---
title: Page Members (Visio)
ms.prod: VISIO
ms.assetid: 655e9dbd-4716-5ee2-37dc-da7e44af4042
---


# Page Members (Visio)
Represents a drawing page, which can be either a foreground page or a background page.

Represents a drawing page, which can be either a foreground page or a background page.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[AfterReplaceShapes](page-afterreplaceshapes-event-visio.md)|Occurs after a shape-replacement operation.|
|[BeforePageDelete](page-beforepagedelete-event-visio.md)|Occurs before a page is deleted.|
|[BeforeReplaceShapes](page-beforereplaceshapes-event-visio.md)|Occurs just before a shape-replacement operation.|
|[BeforeSelectionDelete](page-beforeselectiondelete-event-visio.md)|Occurs before selected objects are deleted.|
|[BeforeShapeDelete](page-beforeshapedelete-event-visio.md)|Occurs before a shape is deleted.|
|[BeforeShapeTextEdit](page-beforeshapetextedit-event-visio.md)|Occurs before a shape is opened for text editing in the user interface.|
|[CalloutRelationshipAdded](page-calloutrelationshipadded-event-visio.md)|Occurs when a new callout relationship is added to the page.|
|[CalloutRelationshipDeleted](page-calloutrelationshipdeleted-event-visio.md)|Occurs when a callout relationship is deleted from the page.|
|[CellChanged](page-cellchanged-event-visio.md)|Occurs after the value changes in a cell in a document.|
|[ConnectionsAdded](page-connectionsadded-event-visio.md)|Occurs after connections have been established between shapes.|
|[ConnectionsDeleted](page-connectionsdeleted-event-visio.md)|Occurs after connections between shapes have been removed.|
|[ContainerRelationshipAdded](page-containerrelationshipadded-event-visio.md)|Occurs when a new container relationship is added to the document.|
|[ContainerRelationshipDeleted](page-containerrelationshipdeleted-event-visio.md)|Occurs when a container relationship is deleted from the document.|
|[ConvertToGroupCanceled](page-converttogroupcanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelConvertToGroup** event.|
|[FormulaChanged](page-formulachanged-event-visio.md)|Occurs after a formula changes in a cell in the object that receives the event.|
|[GroupCanceled](page-groupcanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelGroup** event.|
|[PageChanged](page-pagechanged-event-visio.md)|Occurs after the name of a page, the background page associated with a page, or the page type (foreground or background) changes.|
|[PageDeleteCanceled](page-pagedeletecanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelPageDelete** event.|
|[QueryCancelConvertToGroup](page-querycancelconverttogroup-event-visio.md)|Occurs before the application converts a selection of shapes to a group in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelGroup](page-querycancelgroup-event-visio.md)|Occurs before the application groups a selection of shapes in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelPageDelete](page-querycancelpagedelete-event-visio.md)|Occurs before the application deletes a page in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelReplaceShapes](page-querycancelreplaceshapes-event-visio.md)|Occurs immediately after a shape-replacement operation is requested. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelSelectionDelete](page-querycancelselectiondelete-event-visio.md)|Occurs before the application deletes a selection of shapes in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelUngroup](page-querycancelungroup-event-visio.md)|Occurs before the application ungroups a selection of shapes in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[ReplaceShapesCanceled](page-replaceshapescanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelReplaceShapes** event.|
|[SelectionAdded](page-selectionadded-event-visio.md)|Occurs after one or more shapes are added to a document.|
|[SelectionDeleteCanceled](page-selectiondeletecanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelSelectionDelete** event.|
|[ShapeAdded](page-shapeadded-event-visio.md)|Occurs after one or more shapes are added to a document.|
|[ShapeChanged](page-shapechanged-event-visio.md)|Occurs after a property of a shape that is not stored in a cell is changed in a document.|
|[ShapeDataGraphicChanged](page-shapedatagraphicchanged-event-visio.md)|Occurs after a data graphic is applied to or deleted from a shape.|
|[ShapeExitedTextEdit](page-shapeexitedtextedit-event-visio.md)|Occurs after a shape is no longer open for interactive text editing.|
|[ShapeLinkAdded](page-shapelinkadded-event-visio.md)|Occurs after a shape is linked to a data row.|
|[ShapeLinkDeleted](page-shapelinkdeleted-event-visio.md)|Occurs after the link between a shape and a data row is deleted.|
|[ShapeParentChanged](page-shapeparentchanged-event-visio.md)|Occurs after shapes are grouped or a group is ungrouped.|
|[TextChanged](page-textchanged-event-visio.md)|Occurs after the text of a shape is changed in a document.|
|[UngroupCanceled](page-ungroupcanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelUngroup** event.|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[AddGuide](page-addguide-method-visio.md)|Adds a guide to a drawing page.|
|[AutoConnectMany](page-autoconnectmany-method-visio.md)|Automatically draws multiple connections in the specified directions between the specified shapes. Returns the number of shapes connected.|
|[AutoSizeDrawing](page-autosizedrawing-method-visio.md)|Automatically resizes the drawing page by adding as many printer-paper-sized tiles as necessary to fit all shapes in the drawing onto the page.|
|[AvoidPageBreaks](page-avoidpagebreaks-method-visio.md)|Makes small adjustments to shapes to move them off page breaks.|
|[BoundingBox](page-boundingbox-method-visio.md)|Returns a rectangle that tightly encloses the shapes of a page.|
|[CenterDrawing](page-centerdrawing-method-visio.md)|Centers a page's, master's, or group's shapes with respect to the extent of the page, master, or group. .|
|[CreateSelection](page-createselection-method-visio.md)|Creates various types of  **Selection** objects.|
|[Delete](page-delete-method-visio.md)|Deletes a  **Page** object. Can also renumber remaining pages.|
|[DrawArcByThreePoints](page-drawarcbythreepoints-method-visio.md)|Creates a shape whose path consists of an arc defined by the three points passed as parameters.|
|[DrawBezier](page-drawbezier-method-visio.md)|Creates a shape whose path is defined by the supplied sequence of Bezier control points.|
|[DrawCircularArc](page-drawcirculararc-method-visio.md)|Creates a new shape whose path consists of a circular arc defined by its center, radius, and start and end angles.|
|[DrawLine](page-drawline-method-visio.md)|Adds a line to the  **Shapes** collection of a page.|
|[DrawNURBS](page-drawnurbs-method-visio.md)|Creates a new shape whose path consists of a single NURBS (nonuniform rational B-spline) segment.|
|[DrawOval](page-drawoval-method-visio.md)|Adds an oval (ellipse) to the  **Shapes** collection of a page.|
|[DrawPolyline](page-drawpolyline-method-visio.md)|Creates a shape whose path is a polyline along a given set of points.|
|[DrawQuarterArc](page-drawquarterarc-method-visio.md)|Creates a shape whose path consists of an elliptical arc defined by the two points and the flag passed in as arguments.|
|[DrawRectangle](page-drawrectangle-method-visio.md)|Adds a rectangle to the  **Shapes** collection of a page, master, or group.|
|[DrawSpline](page-drawspline-method-visio.md)|Creates a new shape whose path follows a given sequence of points.|
|[Drop](page-drop-method-visio.md)|Creates one or more new  **Shape** objects by dropping an object onto a receiving object such as a master, drawing page, shape, or group.|
|[DropCallout](page-dropcallout-method-visio.md)|Creates a new callout  **[Shape](shape-object-visio.md)** object on the page near the specified target shape, and associates the callout with the target shape. Returns the callout shape.|
|[DropConnected](page-dropconnected-method-visio.md)|Creates a new  **[Shape](shape-object-visio.md)** object on the page, places the new shape relative to the specified existing target shape, and adds a connector from the existing shape to the new shape. Returns the newly created shape.|
|[DropContainer](page-dropcontainer-method-visio.md)|Creates a new container  **[Shape](shape-object-visio.md)** object on the page, places the container around the specified target shapes, and adds the target shapes to the container. Returns the container shape.|
|[DropIntoList](page-dropintolist-method-visio.md)|Drops the specified object into the specified list at the specified position. Returns the newly dropped shape.|
|[DropLegend](page-droplegend-method-visio.md)|Inserts a data graphics legend on a Microsoft Visio drawing page. Returns the list shape instance specified in the  _OuterList_ parameter.|
|[DropLinked](page-droplinked-method-visio.md)|Returns a new shape on the drawing page linked to data in a data recordset.|
|[DropMany](page-dropmany-method-visio.md)|Creates one or more new  **Shape** objects on a page. It returns an array of the IDs of the **Shape** objects it produces.|
|[DropManyLinkedU](page-dropmanylinkedu-method-visio.md)|Creates multiple new shapes on the drawing page that are linked to multiple data rows in a data recordset. Returns the number of shape instances created and an array of IDs of those shapes.|
|[DropManyU](page-dropmanyu-method-visio.md)|Creates one or more new  **Shape** objects on a page, in a master, or in a group. It returns an array of the IDs of the **Shape** objects it produces.|
|[Duplicate](page-duplicate-method-visio.md)|Duplicates the specified page and returns the new page.|
|[Export](page-export-method-visio.md)|Exports an object from Microsoft Visio to a file format such as .bmp, .dib, .dwg, .dxf, .emf, .emz, .gif, .htm, .jpg, .png, .svg, .svgz, .tif, or .wmf.|
|[GetCallouts](page-getcallouts-method-visio.md)|Returns the list of identifiers of the callout shapes on the page.|
|[GetContainers](page-getcontainers-method-visio.md)|Returns an array of shape identifiers (IDs) of the container shapes on the page.|
|[GetFormulas](page-getformulas-method-visio.md)|Returns the formulas of many cells.|
|[GetFormulasU](page-getformulasu-method-visio.md)|Returns the formulas of many cells.|
|[GetResults](page-getresults-method-visio.md)|Gets the results or formulas of many cells.|
|[GetShapesLinkedToData](page-getshapeslinkedtodata-method-visio.md)|Returns an array of all shapes on the active page linked to data in the specified data recordset.|
|[GetShapesLinkedToDataRow](page-getshapeslinkedtodatarow-method-visio.md)|Returns an array of all shapes on the active page linked to data in the specified data row in the specified data recordset.|
|[GetTheme](page-gettheme-method-visio.md)|Returns a  **Variant** that represents the specified theme component of the specified page.|
|[GetThemeVariant](page-getthemevariant-method-visio.md)|Returns the color, style, and embellishment, if any, of the variant of the theme applied to the specified page.|
|[Import](page-import-method-visio.md)|Imports a file into the current document.|
|[InsertFromFile](page-insertfromfile-method-visio.md)|Adds a linked or embedded object to a page, master, or group.|
|[InsertObject](page-insertobject-method-visio.md)|Adds a new embedded object or ActiveX control to a page, master, or group.|
|[Layout](page-layout-method-visio.md)|Lays out the shapes and/or reroutes the connectors for the page, master, group, or selection.|
|[LayoutChangeDirection](page-layoutchangedirection-method-visio.md)|Revises the layout of a set of connected shapes on the page, by rotating or flipping a connected diagram without rotating or flipping the individual shapes.|
|[LayoutIncremental](page-layoutincremental-method-visio.md)|Makes small adjustments to the position of shapes on the drawing page to better align the shapes or to space them evenly from other shapes.|
|[LinkShapesToDataRows](page-linkshapestodatarows-method-visio.md)|Links multiple rows in the specified data recordset, as specified by their data row IDs, to multiple shapes on the page, and optionally applies the current data graphic to the linked shapes.|
|[OpenDrawWindow](page-opendrawwindow-method-visio.md)|Opens a new drawing window that displays a page.|
|[Paste](page-paste-method-visio.md)|Pastes the contents of the Clipboard into an object.|
|[PasteSpecial](page-pastespecial-method-visio.md)|Inserts the contents of the Clipboard, allowing you to control the format of the pasted information and (optionally) establish a link to the source file (for example, a Microsoft Word document).|
|[PasteToLocation](page-pastetolocation-method-visio.md)|Pastes the shape to the specified location on the page.|
|[Print](page-print-method-visio.md)|Prints the contents of an object to the default printer.|
|[PrintTile](page-printtile-method-visio.md)|Prints a single tile of a drawing page.|
|[ResizeToFitContents](page-resizetofitcontents-method-visio.md)|Resizes the page, or the master's page, to fit tightly around the shapes or master that are on it.|
|[SetFormulas](page-setformulas-method-visio.md)|Sets the formulas of one or more cells.|
|[SetResults](page-setresults-method-visio.md)|Sets the results or formulas of one or more cells.|
|[SetTheme](page-settheme-method-visio.md)|Sets the theme for the specified page.|
|[SetThemeVariant](page-setthemevariant-method-visio.md)|Sets the color, style, and optionally the embellishment of the variant of the theme applied to the specified page.|
|[ShapeIDsToUniqueIDs](page-shapeidstouniqueids-method-visio.md)|Returns an array of unique IDs of shapes on the page, as specified by their shape IDs.|
|[SplitConnector](page-splitconnector-method-visio.md)|Splits the specified connector with the specified shape. Returns the new duplicated connector.|
|[UniqueIDsToShapeIDs](page-uniqueidstoshapeids-method-visio.md)|Returns an array of shape IDs of shapes on the page, as specifed by their unique IDs.|
|[VisualBoundingBox](page-visualboundingbox-method-visio.md)|Returns the bounding rectangle of the virtual container that has all the shapes of the given page.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](page-application-property-visio.md)|Returns the instance of Microsoft Visio that is associated with an object. Read-only.|
|[AutoSize](page-autosize-property-visio.md)|Determines whether Microsoft Visio automatically resizes the drawing page by adding printer-paper-sized sheets, as necessary, to fit the drawing's contents. Read/write.|
|[Background](page-background-property-visio.md)|Determines whether a page is a background page. Read/write.|
|[BackPage](page-backpage-property-visio.md)|Gets or sets the background page of a page. Read/write.|
|[Comments](page-comments-property-visio.md)|Returns a [Comments](comments-object-visio.md) object that represents the collection of all the reviewer comments on the page. Read-only.|
|[Connects](page-connects-property-visio.md)|Returns a  **Connects** collection for a shape, page, or master. Read-only.|
|[Document](page-document-property-visio.md)|Gets the  **Document** object that is associated with an object. Read-only.|
|[EventList](page-eventlist-property-visio.md)|Returns the  **EventList** collection of an object or the **EventList** collection that contains an **Event** object. Read-only.|
|[ID](page-id-property-visio.md)|Gets the ID of an object. Read-only.|
|[Index](page-index-property-visio.md)|Gets or sets the ordinal position of a page in a  **Pages** collection. Read/write.|
|[Layers](page-layers-property-visio.md)|Returns the  **Layers** collection of an object. Read-only.|
|[LayoutRoutePassive](page-layoutroutepassive-property-visio.md)|Determines whether to enable advanced connector routing logic on the page. Read/write.|
|[Name](page-name-property-visio.md)|Specifies the name of an object. Read-only.|
|[NameU](page-nameu-property-visio.md)|Specifies the universal name of a  **Page** object. Read/write.|
|[ObjectType](page-objecttype-property-visio.md)|Returns an object's type. Read-only.|
|[OLEObjects](page-oleobjects-property-visio.md)|Returns the  **OLEObjects** collection of a page. Read-only.|
|[OriginalPage](page-originalpage-property-visio.md)| Returns the **Page** object that represents the original Microsoft Visio drawing page that was marked up on separate markup overlays by reviewers of the drawing. Read-only.|
|[PageSheet](page-pagesheet-property-visio.md)|Returns the page sheet (an object that represents the ShapeSheet spreadsheet) of a page. Read-only.|
|[PersistsEvents](page-persistsevents-property-visio.md)|Indicates whether an object is capable of containing persistent events in its  **EventList** collection. Read-only.|
|[Picture](page-picture-property-visio.md)|Returns a picture that represents an enhanced metafile (EMF) contained in a master, shape, selection, or page. Read-only.|
|[PrintTileCount](page-printtilecount-property-visio.md)|Returns the number of print tiles for a drawing page. Read-only.|
|[ReviewerID](page-reviewerid-property-visio.md)|Gets the reviewer ID associated with the markup overlay. Read-only.|
|[ShapeComments](page-shapecomments-property-visio.md)|Returns a [Comments](comments-object-visio.md) object that represents the collection of all the reviewer comments on all the shapes on the page. Read-only.|
|[Shapes](page-shapes-property-visio.md)|Returns the  **Shapes** collection for a page, master, or group. Read-only.|
|[SpatialSearch](page-spatialsearch-property-visio.md)|Returns a  **Selection** object whose shapes meet certain criteria in relation to a point that is expressed in the coordinate space of a page, master, or group. Read-only.|
|[Stat](page-stat-property-visio.md)|Returns status information for an object. Read-only.|
|[ThemeColors](page-themecolors-property-visio.md)|Gets or sets the current theme colors for the page. Read/write.|
|[ThemeEffects](page-themeeffects-property-visio.md)|Gets or sets the current theme effect for the page. Read/write.|
|[Type](page-type-property-visio.md)|Returns the type of the object. Read-only.|

