---
title: Documents Members (Visio)
ms.prod: VISIO
ms.assetid: 5fff6b50-0883-6c1b-2f2a-2696a0bf5c96
---


# Documents Members (Visio)
 Includes a **Document** object for each open document in a Microsoft Visio instance.

 Includes a **Document** object for each open document in a Microsoft Visio instance.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[AfterDocumentMerge](documents-afterdocumentmerge-event-visio.md)|Occurs when Visio incorporates changes from other users? versions of a document into a merged, co-authored document.|
|[AfterRemoveHiddenInformation](documents-afterremovehiddeninformation-event-visio.md)|Occurs when hidden information is removed from the document.|
|[AfterReplaceShapes](documents-afterreplaceshapes-event-visio.md)|Occurs after a shape-replacement operation.|
|[BeforeDataRecordsetDelete](documents-beforedatarecordsetdelete-event-visio.md)|Occurs before a  **DataRecordset** object is deleted from the **DataRecordsets** collection.|
|[BeforeDocumentClose](documents-beforedocumentclose-event-visio.md)|Occurs before a document is closed.|
|[BeforeDocumentSave](documents-beforedocumentsave-event-visio.md)|Occurs before a document is saved.|
|[BeforeDocumentSaveAs](documents-beforedocumentsaveas-event-visio.md)|Occurs just before a document is saved by using the  **Save As** command.|
|[BeforeMasterDelete](documents-beforemasterdelete-event-visio.md)|Occurs before a master is deleted from a document.|
|[BeforePageDelete](documents-beforepagedelete-event-visio.md)|Occurs before a page is deleted.|
|[BeforeReplaceShapes](documents-beforereplaceshapes-event-visio.md)|Occurs just before a shape-replacement operation.|
|[BeforeSelectionDelete](documents-beforeselectiondelete-event-visio.md)|Occurs before selected objects are deleted.|
|[BeforeShapeDelete](documents-beforeshapedelete-event-visio.md)|Occurs before a shape is deleted.|
|[BeforeShapeTextEdit](documents-beforeshapetextedit-event-visio.md)|Occurs before a shape is opened for text editing in the user interface.|
|[BeforeStyleDelete](documents-beforestyledelete-event-visio.md)|Occurs before a style is deleted.|
|[CalloutRelationshipAdded](documents-calloutrelationshipadded-event-visio.md)|Occurs when a new callout relationship is added to a document.|
|[CalloutRelationshipDeleted](documents-calloutrelationshipdeleted-event-visio.md)|Occurs when a callout relationship is deleted from a document.|
|[CellChanged](documents-cellchanged-event-visio.md)|Occurs after the value changes in a cell in a document.|
|[ConnectionsAdded](documents-connectionsadded-event-visio.md)|Occurs after connections have been established between shapes.|
|[ConnectionsDeleted](documents-connectionsdeleted-event-visio.md)|Occurs after connections between shapes have been removed.|
|[ContainerRelationshipAdded](documents-containerrelationshipadded-event-visio.md)|Occurs when a new container relationship is added to the document.|
|[ContainerRelationshipDeleted](documents-containerrelationshipdeleted-event-visio.md)|Occurs when a container relationship is deleted from the document.|
|[ConvertToGroupCanceled](documents-converttogroupcanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelConvertToGroup** event.|
|[DataRecordsetAdded](documents-datarecordsetadded-event-visio.md)|Occurs when a  **DataRecordset** object is added to a **DataRecordsets** collection.|
|[DataRecordsetChanged](documents-datarecordsetchanged-event-visio.md)|Occurs when a data recordset changes as a result of being refreshed.|
|[DesignModeEntered](documents-designmodeentered-event-visio.md)|Occurs before a document enters design mode.|
|[DocumentChanged](documents-documentchanged-event-visio.md)|Occurs after certain properties of a document are changed.|
|[DocumentCloseCanceled](documents-documentclosecanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelDocumentClose** event.|
|[DocumentCreated](documents-documentcreated-event-visio.md)|Occurs after a document is created.|
|[DocumentOpened](documents-documentopened-event-visio.md)|Occurs after a document is opened.|
|[DocumentSaved](documents-documentsaved-event-visio.md)|Occurs after a document is saved.|
|[DocumentSavedAs](documents-documentsavedas-event-visio.md)|Occurs after a document is saved by using the  **Save As** command.|
|[FormulaChanged](documents-formulachanged-event-visio.md)|Occurs after a formula changes in a cell in the object that receives the event.|
|[GroupCanceled](documents-groupcanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelGroup** event.|
|[MasterAdded](documents-masteradded-event-visio.md)|Occurs after a new master is added to a document.|
|[MasterChanged](documents-masterchanged-event-visio.md)|Occurs after properties of a master are changed and propagated to its instances.|
|[MasterDeleteCanceled](documents-masterdeletecanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelMasterDelete** event.|
|[PageAdded](documents-pageadded-event-visio.md)|Occurs after a new page is added to a document.|
|[PageChanged](documents-pagechanged-event-visio.md)|Occurs after the name of a page, the background page associated with a page, or the page type (foreground or background) changes.|
|[PageDeleteCanceled](documents-pagedeletecanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelPageDelete** event.|
|[QueryCancelConvertToGroup](documents-querycancelconverttogroup-event-visio.md)|Occurs before the application converts a selection of shapes to a group in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelDocumentClose](documents-querycanceldocumentclose-event-visio.md)|Occurs before the application closes a document in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelGroup](documents-querycancelgroup-event-visio.md)|Occurs before the application groups a selection of shapes in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelMasterDelete](documents-querycancelmasterdelete-event-visio.md)|Occurs before the application deletes a master in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelPageDelete](documents-querycancelpagedelete-event-visio.md)|Occurs before the application deletes a page in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelReplaceShapes](documents-querycancelreplaceshapes-event-visio.md)|Occurs immediately after a shape-replacement operation is requested. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelSelectionDelete](documents-querycancelselectiondelete-event-visio.md)|Occurs before the application deletes a selection of shapes in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelStyleDelete](documents-querycancelstyledelete-event-visio.md)|Occurs before the application deletes a style in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelUngroup](documents-querycancelungroup-event-visio.md)|Occurs before the application ungroups a selection of shapes in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[ReplaceShapesCanceled](documents-replaceshapescanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelReplaceShapes** event.|
|[RuleSetValidated](documents-rulesetvalidated-event-visio.md)|Occurs when a rule set is validated.|
|[RunModeEntered](documents-runmodeentered-event-visio.md)|Occurs after a document enters run mode.|
|[SelectionAdded](documents-selectionadded-event-visio.md)|Occurs after one or more shapes are added to a document.|
|[SelectionDeleteCanceled](documents-selectiondeletecanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelSelectionDelete** event.|
|[ShapeAdded](documents-shapeadded-event-visio.md)|Occurs after one or more shapes are added to a document.|
|[ShapeChanged](documents-shapechanged-event-visio.md)|Occurs after a property of a shape that is not stored in a cell is changed in a document.|
|[ShapeDataGraphicChanged](documents-shapedatagraphicchanged-event-visio.md)|Occurs after a data graphic is applied to or deleted from a shape.|
|[ShapeExitedTextEdit](documents-shapeexitedtextedit-event-visio.md)|Occurs after a shape is no longer open for interactive text editing.|
|[ShapeLinkAdded](documents-shapelinkadded-event-visio.md)|Occurs after a shape is linked to a data row.|
|[ShapeLinkDeleted](documents-shapelinkdeleted-event-visio.md)|Occurs after the link between a shape and a data row is deleted.|
|[ShapeParentChanged](documents-shapeparentchanged-event-visio.md)|Occurs after shapes are grouped or a group is ungrouped.|
|[StyleAdded](documents-styleadded-event-visio.md)|Occurs after a new style is added to a document.|
|[StyleChanged](documents-stylechanged-event-visio.md)|Occurs after the name of a style is changed or a change to the style propagates to objects to which the style is applied.|
|[StyleDeleteCanceled](documents-styledeletecanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelStyleDelete** event.|
|[TextChanged](documents-textchanged-event-visio.md)|Occurs after the text of a shape is changed in a document.|
|[UngroupCanceled](documents-ungroupcanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelUngroup** event.|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Add](documents-add-method-visio.md)|Adds a new  **Document** object to the **Documents** collection.|
|[AddEx](documents-addex-method-visio.md)|Adds a new stencil or drawing to the  **Documents** collection, while permitting extra information to be passed in an argument.|
|[CanCheckOut](documents-cancheckout-method-visio.md)|Specifies whether a document can be checked out from a Microsoft SharePoint Server computer.|
|[CheckOut](documents-checkout-method-visio.md)|Marks a specified document as checked out and assigns edit privileges to the current user.|
|[GetNames](documents-getnames-method-visio.md)|Returns the names of all items in a collection.|
|[Open](documents-open-method-visio.md)|Opens an existing file so that it can be edited.|
|[OpenEx](documents-openex-method-visio.md)|Opens an existing Microsoft Visio file, using extra information passed in as an argument.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](documents-application-property-visio.md)|Returns the instance of Microsoft Visio that is associated with an object. Read-only.|
|[Count](documents-count-property-visio.md)|Returns the number of objects in a collection. Read-only.|
|[EventList](documents-eventlist-property-visio.md)|Returns the  **EventList** collection of an object or the **EventList** collection that contains an **Event** object. Read-only.|
|[Item](documents-item-property-visio.md)|Returns an item from a collection. The  **Item** property is the default property for all collections. Read-only.|
|[ItemFromID](documents-itemfromid-property-visio.md)|Returns an item of a collection using the ID of the item. Read-only.|
|[ObjectType](documents-objecttype-property-visio.md)|Returns an object's type. Read-only.|
|[PersistsEvents](documents-persistsevents-property-visio.md)|Indicates whether an object is capable of containing persistent events in its  **EventList** collection. Read-only.|

