---
title: Document Events (Visio)
ms.prod: VISIO
ms.assetid: 76fdbb88-4f3f-4399-8763-bf391ea8fb45
---


# Document Events (Visio)
This object has the following events:

## Events



|**Name**|**Description**|
|:-----|:-----|
|[AfterDocumentMerge](document-afterdocumentmerge-event-visio.md)|Occurs when Visio incorporates changes from other users? versions of a document into a merged, co-authored document.|
|[AfterRemoveHiddenInformation](document-afterremovehiddeninformation-event-visio.md)|Occurs when hidden information is removed from the document.|
|[BeforeDataRecordsetDelete](document-beforedatarecordsetdelete-event-visio.md)|Occurs before a  **DataRecordset** object is deleted from the **DataRecordsets** collection.|
|[BeforeDocumentClose](document-beforedocumentclose-event-visio.md)|Occurs before a document is closed.|
|[BeforeDocumentSave](document-beforedocumentsave-event-visio.md)|Occurs before a document is saved.|
|[BeforeDocumentSaveAs](document-beforedocumentsaveas-event-visio.md)|Occurs just before a document is saved by using the  **Save As** command.|
|[BeforeMasterDelete](document-beforemasterdelete-event-visio.md)|Occurs before a master is deleted from a document.|
|[BeforePageDelete](document-beforepagedelete-event-visio.md)|Occurs before a page is deleted.|
|[BeforeSelectionDelete](document-beforeselectiondelete-event-visio.md)|Occurs before selected objects are deleted.|
|[BeforeShapeTextEdit](document-beforeshapetextedit-event-visio.md)|Occurs before a shape is opened for text editing in the user interface.|
|[BeforeStyleDelete](document-beforestyledelete-event-visio.md)|Occurs before a style is deleted.|
|[ConvertToGroupCanceled](document-converttogroupcanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelConvertToGroup** event.|
|[DataRecordsetAdded](document-datarecordsetadded-event-visio.md)|Occurs when a  **DataRecordset** object is added to a **DataRecordsets** collection.|
|[DesignModeEntered](document-designmodeentered-event-visio.md)|Occurs before a document enters design mode.|
|[DocumentChanged](document-documentchanged-event-visio.md)|Occurs after certain properties of a document are changed.|
|[DocumentCloseCanceled](document-documentclosecanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelDocumentClose** event.|
|[DocumentCreated](document-documentcreated-event-visio.md)|Occurs after a document is created.|
|[DocumentOpened](document-documentopened-event-visio.md)|Occurs after a document is opened.|
|[DocumentSaved](document-documentsaved-event-visio.md)|Occurs after a document is saved.|
|[DocumentSavedAs](document-documentsavedas-event-visio.md)|Occurs after a document is saved by using the  **Save As** command.|
|[GroupCanceled](document-groupcanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelGroup** event.|
|[MasterAdded](document-masteradded-event-visio.md)|Occurs after a new master is added to a document.|
|[MasterChanged](document-masterchanged-event-visio.md)|Occurs after properties of a master are changed and propagated to its instances.|
|[MasterDeleteCanceled](document-masterdeletecanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelMasterDelete** event.|
|[PageAdded](document-pageadded-event-visio.md)|Occurs after a new page is added to a document.|
|[PageChanged](document-pagechanged-event-visio.md)|Occurs after the name of a page, the background page associated with a page, or the page type (foreground or background) changes.|
|[PageDeleteCanceled](document-pagedeletecanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelPageDelete** event.|
|[QueryCancelConvertToGroup](document-querycancelconverttogroup-event-visio.md)|Occurs before the application converts a selection of shapes to a group in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelDocumentClose](document-querycanceldocumentclose-event-visio.md)|Occurs before the application closes a document in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelGroup](document-querycancelgroup-event-visio.md)|Occurs before the application groups a selection of shapes in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelMasterDelete](document-querycancelmasterdelete-event-visio.md)|Occurs before the application deletes a master in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelPageDelete](document-querycancelpagedelete-event-visio.md)|Occurs before the application deletes a page in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelSelectionDelete](document-querycancelselectiondelete-event-visio.md)|Occurs before the application deletes a selection of shapes in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelStyleDelete](document-querycancelstyledelete-event-visio.md)|Occurs before the application deletes a style in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[QueryCancelUngroup](document-querycancelungroup-event-visio.md)|Occurs before the application ungroups a selection of shapes in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.|
|[RuleSetValidated](document-rulesetvalidated-event-visio.md)|Occurs when a rule set is validated.|
|[RunModeEntered](document-runmodeentered-event-visio.md)|Occurs after a document enters run mode.|
|[SelectionDeleteCanceled](document-selectiondeletecanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelSelectionDelete** event.|
|[ShapeAdded](document-shapeadded-event-visio.md)|Occurs after one or more shapes are added to a document.|
|[ShapeDataGraphicChanged](document-shapedatagraphicchanged-event-visio.md)|Occurs after a data graphic is applied to or deleted from a shape.|
|[ShapeExitedTextEdit](document-shapeexitedtextedit-event-visio.md)|Occurs after a shape is no longer open for interactive text editing.|
|[ShapeParentChanged](document-shapeparentchanged-event-visio.md)|Occurs after shapes are grouped or a group is ungrouped.|
|[StyleAdded](document-styleadded-event-visio.md)|Occurs after a new style is added to a document.|
|[StyleChanged](document-stylechanged-event-visio.md)|Occurs after the name of a style is changed or a change to the style propagates to objects to which the style is applied.|
|[StyleDeleteCanceled](document-styledeletecanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelStyleDelete** event.|
|[UngroupCanceled](document-ungroupcanceled-event-visio.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelUngroup** event.|

