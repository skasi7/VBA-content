---
title: Document Members (Visio)
ms.prod: VISIO
ms.assetid: ea706ae9-1287-6f9c-e7de-59167e9f4b09
---


# Document Members (Visio)
Represents a drawing file (.vsd or .vdx), stencil file (.vss or .vsx), or template file (.vst or .vtx) that is open in an instance of Microsoft Visio. A  **Document** object is a member of the **Documents** collection of an **Application** object.

Represents a drawing file (.vsd or .vdx), stencil file (.vss or .vsx), or template file (.vst or .vtx) that is open in an instance of Microsoft Visio. A  **Document** object is a member of the **Documents** collection of an **Application** object.


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

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[AddUndoUnit](document-addundounit-method-visio.md)|Adds an object that supports the  **IOleUndoUnit** or **IVBUndoUnit** interface to the Microsoft Visio undo queue.|
|[BeginUndoScope](document-beginundoscope-method-visio.md)|Starts a transaction with a unique scope ID for an instance of Microsoft Visio.|
|[CanCheckIn](document-cancheckin-method-visio.md)|Specifies whether a document can be checked into a Microsoft SharePoint Server computer.|
|[CanUndoCheckOut](document-canundocheckout-method-visio.md)|Determines whether a Microsoft Visio document is checked out from a Microsoft SharePoint Server site, so that if it is, the check-out can be subsequently undone.|
|[CheckIn](document-checkin-method-visio.md)|Returns a document from a local computer to a Microsoft SharePoint Server computer.|
|[Clean](document-clean-method-visio.md)|Examines, reports, and repairs selected conditions in a document.|
|[ClearCustomMenus](document-clearcustommenus-method-visio.md)|Restores the built-in Microsoft Visio user interface.|
|[ClearCustomToolbars](document-clearcustomtoolbars-method-visio.md)|Restores the built-in Microsoft Visio user interface.|
|[ClearGestureFormatSheet](document-cleargestureformatsheet-method-visio.md)|Clears local formatting in a document's Gesture Format sheet.|
|[Close](document-close-method-visio.md)|Closes a document.|
|[CopyPreviewPicture](document-copypreviewpicture-method-visio.md)|Copies the preview picture from another document into the current document.|
|[DeleteSolutionXMLElement](document-deletesolutionxmlelement-method-visio.md)|Deletes the named SolutionXML element.|
|[Drop](document-drop-method-visio.md)|Creates a new  **Master** object by dropping an object onto a receiving object such as a stencil or document, or the **Masters** or **MasterShortcuts** collection.|
|[EndUndoScope](document-endundoscope-method-visio.md)|Ends or cancels a transaction that has a unique scope.|
|[ExecuteLine](document-executeline-method-visio.md)|Executes a line of Microsoft Visual Basic code.|
|[ExportAsFixedFormat](document-exportasfixedformat-method-visio.md)|Exports a Microsoft Visio document as a file in a fixed format, either PDF or XPS.|
|[FollowHyperlink](document-followhyperlink-method-visio.md)|Navigates to an arbitrary hyperlink.|
|[GetThemeNames](document-getthemenames-method-visio.md)|Returns a locale-specific array of names of themes contained in the document.|
|[GetThemeNamesU](document-getthemenamesu-method-visio.md)|Returns a locale-independent array of names of themes contained in the document.|
|[OpenStencilWindow](document-openstencilwindow-method-visio.md)|Opens a stencil window that shows the masters in the document.|
|[ParseLine](document-parseline-method-visio.md)|Parses a line of Microsoft Visual Basic code.|
|[Print](document-print-method-visio.md)|Prints the contents of an object to the default printer.|
|[PrintOut](document-printout-method-visio.md)|Prints the contents of the active document and provides various printing options.|
|[PurgeUndo](document-purgeundo-method-visio.md)|Empties the Microsoft Visio queue of undo actions.|
|[RemoveHiddenInformation](document-removehiddeninformation-method-visio.md)|Removes hidden information, such as personal information and external data, from a Microsoft Visio document.|
|[RenameCurrentScope](document-renamecurrentscope-method-visio.md)|Renames the top-level open undo scope.|
|[Save](document-save-method-visio.md)|Saves a document.|
|[SaveAs](document-saveas-method-visio.md)|Saves a document and gives it a file name.|
|[SaveAsEx](document-saveasex-method-visio.md)|Saves a document with a file name using extra information passed in an argument.|
|[SetCustomMenus](document-setcustommenus-method-visio.md)|Replaces the current built-in or custom menus of an application or document.|
|[SetCustomToolbars](document-setcustomtoolbars-method-visio.md)|Replaces the current built-in or custom toolbars of an application or document.|
|[UndoCheckOut](document-undocheckout-method-visio.md)|Closes a Microsoft Visio document checked out from a Microsoft SharePoint Server site, deletes the local copy of the document, discarding any changes, undoes the checkout, and then reopens the document.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[AlternateNames](document-alternatenames-property-visio.md)|Gets or sets the alternate names for a document. Read/write.|
|[Application](document-application-property-visio.md)|Returns the instance of Microsoft Visio that is associated with an object. Read-only.|
|[AutoRecover](document-autorecover-property-visio.md)|Determines whether an open document that has unsaved changes is copied when automatic recovery is enabled. Read/write.|
|[BottomMargin](document-bottommargin-property-visio.md)|Specifies the bottom margin when printing the pages in a document. Read/write.|
|[BuildNumberCreated](document-buildnumbercreated-property-visio.md)|Returns the build number of the instance used to create the document. Read-only.|
|[BuildNumberEdited](document-buildnumberedited-property-visio.md)|Returns the build number of the instance last used to edit the document. Read-only.|
|[Category](document-category-property-visio.md)|Gets or sets the value of a document's category, one of the document properties. Read/write.|
|[ClassID](document-classid-property-visio.md)|Returns the class ID string of the container application in which the document is embedded. Read-only.|
|[Colors](document-colors-property-visio.md)|Returns the  **Colors** collection of a **Document** object. Read-only.|
|[Comments](document-comments-property-visio.md)|Returns a [Comments](comments-object-visio.md) object that represents the collection of all the reviewer comments in the document. Read-only.|
|[Company](document-company-property-visio.md)|Gets or sets the name of the company the document belongs to, one of the document's properties. Read/write.|
|[CompatibilityMode](document-compatibilitymode-property-visio.md)|Returns a  **Boolean** that indicates whether the specified document is in compatibility mode. Read-only.|
|[Container](document-container-property-visio.md)|Returns an  **IDispatch** interface on the ActiveX container in which the document is contained or **Nothing** if the document is not in a container. Read-only.|
|[ContainsWorkspaceEx](document-containsworkspaceex-property-visio.md)|Gets or sets whether workspace information is saved with the document. Read/write.|
|[Creator](document-creator-property-visio.md)|Gets or sets the value of a document's author—one of the document's properties. Read/write.|
|[CustomMenus](document-custommenus-property-visio.md)|Gets a  **UIObject** object that represents the current custom menus and accelerators of a **Document** object. Read-only.|
|[CustomMenusFile](document-custommenusfile-property-visio.md)|Gets or sets the name of the file that defines custom menus and accelerators for a  **Document** object. Read/write.|
|[CustomToolbars](document-customtoolbars-property-visio.md)|Gets a  **UIObject** object that represents the current custom toolbars and status bars of a **Document** object. Read-only.|
|[CustomToolbarsFile](document-customtoolbarsfile-property-visio.md)|Returns or sets the name of the file that defines custom toolbars and status bars for a  **Document** object. Read/write.|
|[CustomUI](document-customui-property-visio.md)|Gets or sets the Ribbon XML string that is passed to the document to customize the Microsoft Office Fluent user interface. Read/write.|
|[DataRecordsets](document-datarecordsets-property-visio.md)|Returns the  **DataRecordsets** collection contained in the document. Read-only.|
|[DefaultFillStyle](document-defaultfillstyle-property-visio.md)|Gets or sets the default fill style of a document. Read/write.|
|[DefaultGuideStyle](document-defaultguidestyle-property-visio.md)|Gets or sets the default guide style of a document. Read/write.|
|[DefaultLineStyle](document-defaultlinestyle-property-visio.md)|Gets or sets the default line style of a document. Read/write.|
|[DefaultSavePath](document-defaultsavepath-property-visio.md)|Gets or sets the path to the location where Microsoft Visio saves the current document by default. Read/write.|
|[DefaultStyle](document-defaultstyle-property-visio.md)|Gets the default fill style of a document or sets the default fill, line, and text styles of a document. Read/write.|
|[DefaultTextStyle](document-defaulttextstyle-property-visio.md)|Gets or sets the default text style of a document. Read/write.|
|[Description](document-description-property-visio.md)|Gets or sets the description of a document, one of a document's properties. Read/write.|
|[DiagramServicesEnabled](document-diagramservicesenabled-property-visio.md)|Determines which, if any, diagram services are enabled for the document. Read/write.|
|[DocumentSheet](document-documentsheet-property-visio.md)|Returns a  **Shape** object whose cells represent properties of the document. Read-only.|
|[DynamicGridEnabled](document-dynamicgridenabled-property-visio.md)|Determines whether the dynamic grid is enabled. Read/write.|
|[EditorCount](document-editorcount-property-visio.md)|Returns a  **Long** that represents the number of editors of a co-authored document. Read-only.|
|[EmailRoutingData](document-emailroutingdata-property-visio.md)|Returns e-mail routing data for a document. Read-only. |
|[EventList](document-eventlist-property-visio.md)|Returns the  **EventList** collection of an object or the **EventList** collection that contains an **Event** object. Read-only.|
|[Fonts](document-fonts-property-visio.md)|Returns the  **Fonts** collection of a **Document** object. Read-only.|
|[FooterCenter](document-footercenter-property-visio.md)|Gets or sets the text string that appears in the center portion of a document's footer. Read/write.|
|[FooterLeft](document-footerleft-property-visio.md)|Gets or sets the text string that appears on the left side of a document's footer. Read/write.|
|[FooterMargin](document-footermargin-property-visio.md)|Gets or sets the margin of a document's footer. Read/write.|
|[FooterRight](document-footerright-property-visio.md)|Gets or sets the text string that appears in the right portion of a document's footer. Read/write.|
|[FullBuildNumberCreated](document-fullbuildnumbercreated-property-visio.md)|Returns the full build number of the instance used to create the document. Read-only.|
|[FullBuildNumberEdited](document-fullbuildnumberedited-property-visio.md)|Returns the full build number of the instance last used to edit the document. Read-only. |
|[FullName](document-fullname-property-visio.md)|Returns the name of a document, including the drive and path. Read-only.|
|[GestureFormatSheet](document-gestureformatsheet-property-visio.md)|Returns a reference to a document's Gesture Format sheet, which contains the line, fill, and text formatting that is applied to shapes drawn on the page. Read-only.|
|[GlueEnabled](document-glueenabled-property-visio.md)|Determines whether glue is enabled in the document. Read/write.|
|[GlueSettings](document-gluesettings-property-visio.md)|Determines the objects that shapes glue to when glue is enabled in the document. Read/write.|
|[HeaderCenter](document-headercenter-property-visio.md)|Contains the text string that appears in the center portion of a document's header. Read/write.|
|[HeaderFooterColor](document-headerfootercolor-property-visio.md)|Specifies the color of the header and footer text. Read/write.|
|[HeaderFooterFont](document-headerfooterfont-property-visio.md)|Specifies the font used for the header and footer text. Read/write.|
|[HeaderLeft](document-headerleft-property-visio.md)|Gets or sets the text string that appears in the left portion of a document's header. Read/write.|
|[HeaderMargin](document-headermargin-property-visio.md)|Gets or sets the margin of a document's header. Read/write.|
|[HeaderRight](document-headerright-property-visio.md)|Gets or sets the text string that appears in the right portion of a document's header. Read/write.|
|[HyperlinkBase](document-hyperlinkbase-property-visio.md)|Gets or sets the value of the  **Hyperlink base** box in a document's **Properties** dialog box (click the **File** tab, click **Info**, click  **Properties**, and then click  **Advanced Properties**). Read/write.|
|[ID](document-id-property-visio.md)|Gets the ID of an object. Read-only.|
|[Index](document-index-property-visio.md)|Gets the ordinal position of a  **Document** object in the **Documents** collection. Read-only.|
|[InPlace](document-inplace-property-visio.md)|Specifies whether a document is open in place, or whether a document is being viewed through a window that is open in place. Read-only.|
|[Keywords](document-keywords-property-visio.md)|Gets or sets the value of the  **Keywords** box in a document's **Properties** dialog box. Read/write.|
|[Language](document-language-property-visio.md)|Represents the language ID of the version of the Microsoft Visio instance represented by the parent object. Read/write.|
|[LeftMargin](document-leftmargin-property-visio.md)|Specifies the left margin, which is used when printing. Read/write.|
|[MacrosEnabled](document-macrosenabled-property-visio.md)|Specifies whether you can execute macros and process events in a document's Microsoft Visual Basic for Applications (VBA) project. Read-only.|
|[Manager](document-manager-property-visio.md)|Gets or sets the value of the  **Manager** box in a document's **Properties** dialog box. Read/write.|
|[Masters](document-masters-property-visio.md)|Returns the  **Masters** collection for a document's stencil. Read-only.|
|[MasterShortcuts](document-mastershortcuts-property-visio.md)|Returns the  **MasterShortcuts** collection for a document stencil. Read-only.|
|[Mode](document-mode-property-visio.md)|Determines whether a document is in run mode or design mode. Read/write.|
|[Name](document-name-property-visio.md)|Specifies the name of an object. Read-only.|
|[ObjectType](document-objecttype-property-visio.md)|Returns an object's type. Read-only.|
|[OLEObjects](document-oleobjects-property-visio.md)|Returns the  **OLEObjects** collection of a document. Read-only.|
|[Pages](document-pages-property-visio.md)|Returns the  **Pages** collection for a document. Read-only.|
|[PaperHeight](document-paperheight-property-visio.md)|Returns the height of a document's printed page. Read-only.|
|[PaperSize](document-papersize-property-visio.md)|Gets or sets the paper size of a document. Read/write.|
|[PaperWidth](document-paperwidth-property-visio.md)|Returns the width of a document's printed page. Read-only.|
|[Path](document-path-property-visio.md)|Returns the drive and folder path of the Microsoft Visio document. Read-only.|
|[PersistsEvents](document-persistsevents-property-visio.md)|Indicates whether an object is capable of containing persistent events in its  **EventList** collection. Read-only.|
|[PreviewPicture](document-previewpicture-property-visio.md)|Gets or sets the preview picture shown in the  **Open** dialog box and when you click the **File** tab, and then click **New**. Read/write.|
|[PrintCenteredH](document-printcenteredh-property-visio.md)|Indicates whether drawings are centered between the left and right edges of the paper when printed. Read/write.|
|[PrintCenteredV](document-printcenteredv-property-visio.md)|Indicates whether drawings are centered between the top and bottom edges of the paper when printed. Read/write.|
|[Printer](document-printer-property-visio.md)|Specifies the name of the printer to use when printing the document. Read/write.|
|[PrintFitOnPages](document-printfitonpages-property-visio.md)|Indicates whether drawings in a document are printed on a specified number of sheets across and down. Read/write.|
|[PrintLandscape](document-printlandscape-property-visio.md)|Indicates whether a document's drawings are printed in landscape or portrait orientation. Read/write.|
|[PrintPagesAcross](document-printpagesacross-property-visio.md)|Indicates the number of sheets of paper on which a drawing is printed horizontally. Read/write.|
|[PrintPagesDown](document-printpagesdown-property-visio.md)|Gets or sets the number of sheets of paper on which a drawing is printed vertically. Read/write.|
|[PrintScale](document-printscale-property-visio.md)|Gets or sets how much drawings are reduced or enlarged when printed. Read/write.|
|[ProgID](document-progid-property-visio.md)|Returns the programmatic identifier of a shape that represents an ActiveX control, an embedded object, or linked object. Read-only.|
|[Protection](document-protection-property-visio.md)|Determines how a document is protected from user customization. Read/write.|
|[ReadOnly](document-readonly-property-visio.md)|Indicates whether a file is open as read-only. Read-only.|
|[RemovePersonalInformation](document-removepersonalinformation-property-visio.md)|Determines if personal information about a file is saved when the user saves the file in Microsoft Visio. Read/write.|
|[RightMargin](document-rightmargin-property-visio.md)|Specifies the right margin, which is used when printing. Read/write.|
|[Saved](document-saved-property-visio.md)|Indicates whether a document has any unsaved changes. Read/write.|
|[SavePreviewMode](document-savepreviewmode-property-visio.md)|Determines whether and how a preview picture is saved in a file. Read/write.|
|[ServerPublishOptions](document-serverpublishoptions-property-visio.md)|Returns a  **[ServerPublishOptions](serverpublishoptions-object-visio.md)** object that you can use to specify the settings to apply when you save a document as a Web drawing (as a .vdw file), then publish and use it on a Microsoft SharePoint site in conjunction with Visio Services. Read-only.|
|[SharedWorkspace](document-sharedworkspace-property-visio.md)|Returns a Microsoft Office  **SharedWorkspace** object that provides access to the Office Document Workspace object model. Read-only.|
|[SnapAngles](document-snapangles-property-visio.md)|Determines the degree of the angle that is drawn when isometric angle lines is chosen as a shape extension option. Read/write.|
|[SnapEnabled](document-snapenabled-property-visio.md)|Determines whether snap is active in the document. Read/write.|
|[SnapExtensions](document-snapextensions-property-visio.md)|Determines the shape extensions that are active in a document. Read/write.|
|[SnapSettings](document-snapsettings-property-visio.md)|Determines the objects that shapes snap to when snap is active in the document. Read/write.|
|[SolutionXMLElement](document-solutionxmlelement-property-visio.md)|Contains solution-specific, well-formed XML data stored with a document. Read/write.|
|[SolutionXMLElementCount](document-solutionxmlelementcount-property-visio.md)|Returns the number of SolutionXML elements in a document. Read-only.|
|[SolutionXMLElementExists](document-solutionxmlelementexists-property-visio.md)|Indicates whether a named SolutionXML element exists in the document. Read-only.|
|[SolutionXMLElementName](document-solutionxmlelementname-property-visio.md)|Returns the name of the SolutionXML element. Read-only.|
|[Stat](document-stat-property-visio.md)|Returns status information for an object. Read-only.|
|[Styles](document-styles-property-visio.md)|Returns the  **Styles** collection for a document. Read-only.|
|[Subject](document-subject-property-visio.md)|Gets or sets the value of the  **Subject** field in a document's properties. Read/write.|
|[Sync](document-sync-property-visio.md)|Returns a Microsoft Office  **Sync** object that provides information about the status of the active document in a shared workspace and the ability to perform a set of actions. Read-only.|
|[Template](document-template-property-visio.md)|Returns the name of the template from which the document was created. Read-only.|
|[Time](document-time-property-visio.md)|Returns the most recently recorded date and time. Read-only.|
|[TimeCreated](document-timecreated-property-visio.md)|Returns the date and time the document was created. Read-only.|
|[TimeEdited](document-timeedited-property-visio.md)|Returns the date and time the document was last edited. Read-only.|
|[TimePrinted](document-timeprinted-property-visio.md)|Returns the date and time the document was last printed. Read-only.|
|[TimeSaved](document-timesaved-property-visio.md)|Returns the date and time the document was last saved. Read-only.|
|[Title](document-title-property-visio.md)|Gets or sets the value of the  **Title** field in a document's properties. Read/write.|
|[TopMargin](document-topmargin-property-visio.md)|Specifies the top margin when printing a document. Read/write.|
|[Type](document-type-property-visio.md)|Returns the type of the  **Document** object. Read-only.|
|[UndoEnabled](document-undoenabled-property-visio.md)|Determines whether undo information is maintained in memory. Read/write.|
|[UserCustomUI](document-usercustomui-property-visio.md)|Gets or sets the Ribbon XML string that is passed to the document to customize the  **Quick Access** toolbar or the Ribbon. Read/write.|
|[Validation](document-validation-property-visio.md)|Returns the  **[Validation](validation-object-visio.md)** object that is associated with the document. Read-only.|
|[VBProject](document-vbproject-property-visio.md)|Returns an automation object that you can use to control the Microsoft Visual Basic for Applications (VBA) project of the document. Read-only.|
|[VBProjectData](document-vbprojectdata-property-visio.md)|Returns the Microsoft Visual Basic project data stored with a document. Read-only.|
|[Version](document-version-property-visio.md)|Gets the version of a saved document or sets the version in which to save a document. Read/write.|
|[ZoomBehavior](document-zoombehavior-property-visio.md)|Determines the zoom behavior for a Microsoft Visio document or window. Read/write.|
|[Permission](document-permission-property-visio.md)||

