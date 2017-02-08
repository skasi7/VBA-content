---
title: Document Object (Visio)
keywords: vis_sdr.chm10080
f1_keywords:
- vis_sdr.chm10080
ms.prod: VISIO
api_name:
- Visio.Document
ms.assetid: 21640062-13a2-a2b2-7c61-7e707671207c
---


# Document Object (Visio)

Represents a drawing file (.vsd or .vdx), stencil file (.vss or .vsx), or template file (.vst or .vtx) that is open in an instance of Microsoft Visio. A  **Document** object is a member of the **Documents** collection of an **Application** object.


## Remarks

The default property of a  **Document** object is **Name**.

Use the  **Open** method of a **Documents** collection to open an existing document.

Use the  **Add** method of a **Documents** collection to create a new document.

Use the  **ActiveDocument** property of an **Application** object to retrieve the active document in an instance.

Use the  **Pages**, **Masters**, and **Styles** properties of a **Document** object to retrieve **Page**, **Master**, and **Style** objects, respectively.


 **Note**  

Use the  **CustomMenus** or **CustomToolbars** properties of a **Document** object to access the custom menus or toolbars.


 **Note**   The Microsoft Visual Basic for Applications (VBA) project of every Visio document also has a class module called **ThisDocument**. When you reference the **ThisDocument** module from code in a VBA project, it returns a reference to the project's **Document** object. For example, the code in a document's project can display the name of the project's document in a **message** box with this statement:




```
    MsgBox ThisDocument.Name
```

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this object maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVDocument**
    

## Events



|**Name**|
|:-----|
|[AfterDocumentMerge](http://msdn.microsoft.com/library/document-afterdocumentmerge-event-visio%28Office.15%29.aspx)|
|[AfterRemoveHiddenInformation](http://msdn.microsoft.com/library/document-afterremovehiddeninformation-event-visio%28Office.15%29.aspx)|
|[BeforeDataRecordsetDelete](http://msdn.microsoft.com/library/document-beforedatarecordsetdelete-event-visio%28Office.15%29.aspx)|
|[BeforeDocumentClose](http://msdn.microsoft.com/library/document-beforedocumentclose-event-visio%28Office.15%29.aspx)|
|[BeforeDocumentSave](http://msdn.microsoft.com/library/document-beforedocumentsave-event-visio%28Office.15%29.aspx)|
|[BeforeDocumentSaveAs](http://msdn.microsoft.com/library/document-beforedocumentsaveas-event-visio%28Office.15%29.aspx)|
|[BeforeMasterDelete](http://msdn.microsoft.com/library/document-beforemasterdelete-event-visio%28Office.15%29.aspx)|
|[BeforePageDelete](http://msdn.microsoft.com/library/document-beforepagedelete-event-visio%28Office.15%29.aspx)|
|[BeforeSelectionDelete](http://msdn.microsoft.com/library/document-beforeselectiondelete-event-visio%28Office.15%29.aspx)|
|[BeforeShapeTextEdit](http://msdn.microsoft.com/library/document-beforeshapetextedit-event-visio%28Office.15%29.aspx)|
|[BeforeStyleDelete](http://msdn.microsoft.com/library/document-beforestyledelete-event-visio%28Office.15%29.aspx)|
|[ConvertToGroupCanceled](http://msdn.microsoft.com/library/document-converttogroupcanceled-event-visio%28Office.15%29.aspx)|
|[DataRecordsetAdded](http://msdn.microsoft.com/library/document-datarecordsetadded-event-visio%28Office.15%29.aspx)|
|[DesignModeEntered](http://msdn.microsoft.com/library/document-designmodeentered-event-visio%28Office.15%29.aspx)|
|[DocumentChanged](http://msdn.microsoft.com/library/document-documentchanged-event-visio%28Office.15%29.aspx)|
|[DocumentCloseCanceled](http://msdn.microsoft.com/library/document-documentclosecanceled-event-visio%28Office.15%29.aspx)|
|[DocumentCreated](http://msdn.microsoft.com/library/document-documentcreated-event-visio%28Office.15%29.aspx)|
|[DocumentOpened](http://msdn.microsoft.com/library/document-documentopened-event-visio%28Office.15%29.aspx)|
|[DocumentSaved](http://msdn.microsoft.com/library/document-documentsaved-event-visio%28Office.15%29.aspx)|
|[DocumentSavedAs](http://msdn.microsoft.com/library/document-documentsavedas-event-visio%28Office.15%29.aspx)|
|[GroupCanceled](http://msdn.microsoft.com/library/document-groupcanceled-event-visio%28Office.15%29.aspx)|
|[MasterAdded](http://msdn.microsoft.com/library/document-masteradded-event-visio%28Office.15%29.aspx)|
|[MasterChanged](http://msdn.microsoft.com/library/document-masterchanged-event-visio%28Office.15%29.aspx)|
|[MasterDeleteCanceled](http://msdn.microsoft.com/library/document-masterdeletecanceled-event-visio%28Office.15%29.aspx)|
|[PageAdded](http://msdn.microsoft.com/library/document-pageadded-event-visio%28Office.15%29.aspx)|
|[PageChanged](http://msdn.microsoft.com/library/document-pagechanged-event-visio%28Office.15%29.aspx)|
|[PageDeleteCanceled](http://msdn.microsoft.com/library/document-pagedeletecanceled-event-visio%28Office.15%29.aspx)|
|[QueryCancelConvertToGroup](http://msdn.microsoft.com/library/document-querycancelconverttogroup-event-visio%28Office.15%29.aspx)|
|[QueryCancelDocumentClose](http://msdn.microsoft.com/library/document-querycanceldocumentclose-event-visio%28Office.15%29.aspx)|
|[QueryCancelGroup](http://msdn.microsoft.com/library/document-querycancelgroup-event-visio%28Office.15%29.aspx)|
|[QueryCancelMasterDelete](http://msdn.microsoft.com/library/document-querycancelmasterdelete-event-visio%28Office.15%29.aspx)|
|[QueryCancelPageDelete](http://msdn.microsoft.com/library/document-querycancelpagedelete-event-visio%28Office.15%29.aspx)|
|[QueryCancelSelectionDelete](http://msdn.microsoft.com/library/document-querycancelselectiondelete-event-visio%28Office.15%29.aspx)|
|[QueryCancelStyleDelete](http://msdn.microsoft.com/library/document-querycancelstyledelete-event-visio%28Office.15%29.aspx)|
|[QueryCancelUngroup](http://msdn.microsoft.com/library/document-querycancelungroup-event-visio%28Office.15%29.aspx)|
|[RuleSetValidated](http://msdn.microsoft.com/library/document-rulesetvalidated-event-visio%28Office.15%29.aspx)|
|[RunModeEntered](http://msdn.microsoft.com/library/document-runmodeentered-event-visio%28Office.15%29.aspx)|
|[SelectionDeleteCanceled](http://msdn.microsoft.com/library/document-selectiondeletecanceled-event-visio%28Office.15%29.aspx)|
|[ShapeAdded](http://msdn.microsoft.com/library/document-shapeadded-event-visio%28Office.15%29.aspx)|
|[ShapeDataGraphicChanged](http://msdn.microsoft.com/library/document-shapedatagraphicchanged-event-visio%28Office.15%29.aspx)|
|[ShapeExitedTextEdit](http://msdn.microsoft.com/library/document-shapeexitedtextedit-event-visio%28Office.15%29.aspx)|
|[ShapeParentChanged](http://msdn.microsoft.com/library/document-shapeparentchanged-event-visio%28Office.15%29.aspx)|
|[StyleAdded](http://msdn.microsoft.com/library/document-styleadded-event-visio%28Office.15%29.aspx)|
|[StyleChanged](http://msdn.microsoft.com/library/document-stylechanged-event-visio%28Office.15%29.aspx)|
|[StyleDeleteCanceled](http://msdn.microsoft.com/library/document-styledeletecanceled-event-visio%28Office.15%29.aspx)|
|[UngroupCanceled](http://msdn.microsoft.com/library/document-ungroupcanceled-event-visio%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[AddUndoUnit](http://msdn.microsoft.com/library/document-addundounit-method-visio%28Office.15%29.aspx)|
|[BeginUndoScope](http://msdn.microsoft.com/library/document-beginundoscope-method-visio%28Office.15%29.aspx)|
|[CanCheckIn](http://msdn.microsoft.com/library/document-cancheckin-method-visio%28Office.15%29.aspx)|
|[CanUndoCheckOut](http://msdn.microsoft.com/library/document-canundocheckout-method-visio%28Office.15%29.aspx)|
|[CheckIn](http://msdn.microsoft.com/library/document-checkin-method-visio%28Office.15%29.aspx)|
|[Clean](http://msdn.microsoft.com/library/document-clean-method-visio%28Office.15%29.aspx)|
|[ClearCustomMenus](http://msdn.microsoft.com/library/document-clearcustommenus-method-visio%28Office.15%29.aspx)|
|[ClearCustomToolbars](http://msdn.microsoft.com/library/document-clearcustomtoolbars-method-visio%28Office.15%29.aspx)|
|[ClearGestureFormatSheet](http://msdn.microsoft.com/library/document-cleargestureformatsheet-method-visio%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/document-close-method-visio%28Office.15%29.aspx)|
|[CopyPreviewPicture](http://msdn.microsoft.com/library/document-copypreviewpicture-method-visio%28Office.15%29.aspx)|
|[DeleteSolutionXMLElement](http://msdn.microsoft.com/library/document-deletesolutionxmlelement-method-visio%28Office.15%29.aspx)|
|[Drop](http://msdn.microsoft.com/library/document-drop-method-visio%28Office.15%29.aspx)|
|[EndUndoScope](http://msdn.microsoft.com/library/document-endundoscope-method-visio%28Office.15%29.aspx)|
|[ExecuteLine](http://msdn.microsoft.com/library/document-executeline-method-visio%28Office.15%29.aspx)|
|[ExportAsFixedFormat](http://msdn.microsoft.com/library/document-exportasfixedformat-method-visio%28Office.15%29.aspx)|
|[FollowHyperlink](http://msdn.microsoft.com/library/document-followhyperlink-method-visio%28Office.15%29.aspx)|
|[GetThemeNames](http://msdn.microsoft.com/library/document-getthemenames-method-visio%28Office.15%29.aspx)|
|[GetThemeNamesU](http://msdn.microsoft.com/library/document-getthemenamesu-method-visio%28Office.15%29.aspx)|
|[OpenStencilWindow](http://msdn.microsoft.com/library/document-openstencilwindow-method-visio%28Office.15%29.aspx)|
|[ParseLine](http://msdn.microsoft.com/library/document-parseline-method-visio%28Office.15%29.aspx)|
|[Print](http://msdn.microsoft.com/library/document-print-method-visio%28Office.15%29.aspx)|
|[PrintOut](http://msdn.microsoft.com/library/document-printout-method-visio%28Office.15%29.aspx)|
|[PurgeUndo](http://msdn.microsoft.com/library/document-purgeundo-method-visio%28Office.15%29.aspx)|
|[RemoveHiddenInformation](http://msdn.microsoft.com/library/document-removehiddeninformation-method-visio%28Office.15%29.aspx)|
|[RenameCurrentScope](http://msdn.microsoft.com/library/document-renamecurrentscope-method-visio%28Office.15%29.aspx)|
|[Save](http://msdn.microsoft.com/library/document-save-method-visio%28Office.15%29.aspx)|
|[SaveAs](http://msdn.microsoft.com/library/document-saveas-method-visio%28Office.15%29.aspx)|
|[SaveAsEx](http://msdn.microsoft.com/library/document-saveasex-method-visio%28Office.15%29.aspx)|
|[SetCustomMenus](http://msdn.microsoft.com/library/document-setcustommenus-method-visio%28Office.15%29.aspx)|
|[SetCustomToolbars](http://msdn.microsoft.com/library/document-setcustomtoolbars-method-visio%28Office.15%29.aspx)|
|[UndoCheckOut](http://msdn.microsoft.com/library/document-undocheckout-method-visio%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AlternateNames](http://msdn.microsoft.com/library/document-alternatenames-property-visio%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/document-application-property-visio%28Office.15%29.aspx)|
|[AutoRecover](http://msdn.microsoft.com/library/document-autorecover-property-visio%28Office.15%29.aspx)|
|[BottomMargin](http://msdn.microsoft.com/library/document-bottommargin-property-visio%28Office.15%29.aspx)|
|[BuildNumberCreated](http://msdn.microsoft.com/library/document-buildnumbercreated-property-visio%28Office.15%29.aspx)|
|[BuildNumberEdited](http://msdn.microsoft.com/library/document-buildnumberedited-property-visio%28Office.15%29.aspx)|
|[Category](http://msdn.microsoft.com/library/document-category-property-visio%28Office.15%29.aspx)|
|[ClassID](http://msdn.microsoft.com/library/document-classid-property-visio%28Office.15%29.aspx)|
|[Colors](http://msdn.microsoft.com/library/document-colors-property-visio%28Office.15%29.aspx)|
|[Comments](http://msdn.microsoft.com/library/document-comments-property-visio%28Office.15%29.aspx)|
|[Company](http://msdn.microsoft.com/library/document-company-property-visio%28Office.15%29.aspx)|
|[CompatibilityMode](http://msdn.microsoft.com/library/document-compatibilitymode-property-visio%28Office.15%29.aspx)|
|[Container](http://msdn.microsoft.com/library/document-container-property-visio%28Office.15%29.aspx)|
|[ContainsWorkspaceEx](http://msdn.microsoft.com/library/document-containsworkspaceex-property-visio%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/document-creator-property-visio%28Office.15%29.aspx)|
|[CustomMenus](http://msdn.microsoft.com/library/document-custommenus-property-visio%28Office.15%29.aspx)|
|[CustomMenusFile](http://msdn.microsoft.com/library/document-custommenusfile-property-visio%28Office.15%29.aspx)|
|[CustomToolbars](http://msdn.microsoft.com/library/document-customtoolbars-property-visio%28Office.15%29.aspx)|
|[CustomToolbarsFile](http://msdn.microsoft.com/library/document-customtoolbarsfile-property-visio%28Office.15%29.aspx)|
|[CustomUI](http://msdn.microsoft.com/library/document-customui-property-visio%28Office.15%29.aspx)|
|[DataRecordsets](http://msdn.microsoft.com/library/document-datarecordsets-property-visio%28Office.15%29.aspx)|
|[DefaultFillStyle](http://msdn.microsoft.com/library/document-defaultfillstyle-property-visio%28Office.15%29.aspx)|
|[DefaultGuideStyle](http://msdn.microsoft.com/library/document-defaultguidestyle-property-visio%28Office.15%29.aspx)|
|[DefaultLineStyle](http://msdn.microsoft.com/library/document-defaultlinestyle-property-visio%28Office.15%29.aspx)|
|[DefaultSavePath](http://msdn.microsoft.com/library/document-defaultsavepath-property-visio%28Office.15%29.aspx)|
|[DefaultStyle](http://msdn.microsoft.com/library/document-defaultstyle-property-visio%28Office.15%29.aspx)|
|[DefaultTextStyle](http://msdn.microsoft.com/library/document-defaulttextstyle-property-visio%28Office.15%29.aspx)|
|[Description](http://msdn.microsoft.com/library/document-description-property-visio%28Office.15%29.aspx)|
|[DiagramServicesEnabled](http://msdn.microsoft.com/library/document-diagramservicesenabled-property-visio%28Office.15%29.aspx)|
|[DocumentSheet](http://msdn.microsoft.com/library/document-documentsheet-property-visio%28Office.15%29.aspx)|
|[DynamicGridEnabled](http://msdn.microsoft.com/library/document-dynamicgridenabled-property-visio%28Office.15%29.aspx)|
|[EditorCount](http://msdn.microsoft.com/library/document-editorcount-property-visio%28Office.15%29.aspx)|
|[EmailRoutingData](http://msdn.microsoft.com/library/document-emailroutingdata-property-visio%28Office.15%29.aspx)|
|[EventList](http://msdn.microsoft.com/library/document-eventlist-property-visio%28Office.15%29.aspx)|
|[Fonts](http://msdn.microsoft.com/library/document-fonts-property-visio%28Office.15%29.aspx)|
|[FooterCenter](http://msdn.microsoft.com/library/document-footercenter-property-visio%28Office.15%29.aspx)|
|[FooterLeft](http://msdn.microsoft.com/library/document-footerleft-property-visio%28Office.15%29.aspx)|
|[FooterMargin](http://msdn.microsoft.com/library/document-footermargin-property-visio%28Office.15%29.aspx)|
|[FooterRight](http://msdn.microsoft.com/library/document-footerright-property-visio%28Office.15%29.aspx)|
|[FullBuildNumberCreated](http://msdn.microsoft.com/library/document-fullbuildnumbercreated-property-visio%28Office.15%29.aspx)|
|[FullBuildNumberEdited](http://msdn.microsoft.com/library/document-fullbuildnumberedited-property-visio%28Office.15%29.aspx)|
|[FullName](http://msdn.microsoft.com/library/document-fullname-property-visio%28Office.15%29.aspx)|
|[GestureFormatSheet](http://msdn.microsoft.com/library/document-gestureformatsheet-property-visio%28Office.15%29.aspx)|
|[GlueEnabled](http://msdn.microsoft.com/library/document-glueenabled-property-visio%28Office.15%29.aspx)|
|[GlueSettings](http://msdn.microsoft.com/library/document-gluesettings-property-visio%28Office.15%29.aspx)|
|[HeaderCenter](http://msdn.microsoft.com/library/document-headercenter-property-visio%28Office.15%29.aspx)|
|[HeaderFooterColor](http://msdn.microsoft.com/library/document-headerfootercolor-property-visio%28Office.15%29.aspx)|
|[HeaderFooterFont](http://msdn.microsoft.com/library/document-headerfooterfont-property-visio%28Office.15%29.aspx)|
|[HeaderLeft](http://msdn.microsoft.com/library/document-headerleft-property-visio%28Office.15%29.aspx)|
|[HeaderMargin](http://msdn.microsoft.com/library/document-headermargin-property-visio%28Office.15%29.aspx)|
|[HeaderRight](http://msdn.microsoft.com/library/document-headerright-property-visio%28Office.15%29.aspx)|
|[HyperlinkBase](http://msdn.microsoft.com/library/document-hyperlinkbase-property-visio%28Office.15%29.aspx)|
|[ID](http://msdn.microsoft.com/library/document-id-property-visio%28Office.15%29.aspx)|
|[Index](http://msdn.microsoft.com/library/document-index-property-visio%28Office.15%29.aspx)|
|[InPlace](http://msdn.microsoft.com/library/document-inplace-property-visio%28Office.15%29.aspx)|
|[Keywords](http://msdn.microsoft.com/library/document-keywords-property-visio%28Office.15%29.aspx)|
|[Language](http://msdn.microsoft.com/library/document-language-property-visio%28Office.15%29.aspx)|
|[LeftMargin](http://msdn.microsoft.com/library/document-leftmargin-property-visio%28Office.15%29.aspx)|
|[MacrosEnabled](http://msdn.microsoft.com/library/document-macrosenabled-property-visio%28Office.15%29.aspx)|
|[Manager](http://msdn.microsoft.com/library/document-manager-property-visio%28Office.15%29.aspx)|
|[Masters](http://msdn.microsoft.com/library/document-masters-property-visio%28Office.15%29.aspx)|
|[MasterShortcuts](http://msdn.microsoft.com/library/document-mastershortcuts-property-visio%28Office.15%29.aspx)|
|[Mode](http://msdn.microsoft.com/library/document-mode-property-visio%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/document-name-property-visio%28Office.15%29.aspx)|
|[ObjectType](http://msdn.microsoft.com/library/document-objecttype-property-visio%28Office.15%29.aspx)|
|[OLEObjects](http://msdn.microsoft.com/library/document-oleobjects-property-visio%28Office.15%29.aspx)|
|[Pages](http://msdn.microsoft.com/library/document-pages-property-visio%28Office.15%29.aspx)|
|[PaperHeight](http://msdn.microsoft.com/library/document-paperheight-property-visio%28Office.15%29.aspx)|
|[PaperSize](http://msdn.microsoft.com/library/document-papersize-property-visio%28Office.15%29.aspx)|
|[PaperWidth](http://msdn.microsoft.com/library/document-paperwidth-property-visio%28Office.15%29.aspx)|
|[Path](http://msdn.microsoft.com/library/document-path-property-visio%28Office.15%29.aspx)|
|[PersistsEvents](http://msdn.microsoft.com/library/document-persistsevents-property-visio%28Office.15%29.aspx)|
|[PreviewPicture](http://msdn.microsoft.com/library/document-previewpicture-property-visio%28Office.15%29.aspx)|
|[PrintCenteredH](http://msdn.microsoft.com/library/document-printcenteredh-property-visio%28Office.15%29.aspx)|
|[PrintCenteredV](http://msdn.microsoft.com/library/document-printcenteredv-property-visio%28Office.15%29.aspx)|
|[Printer](http://msdn.microsoft.com/library/document-printer-property-visio%28Office.15%29.aspx)|
|[PrintFitOnPages](http://msdn.microsoft.com/library/document-printfitonpages-property-visio%28Office.15%29.aspx)|
|[PrintLandscape](http://msdn.microsoft.com/library/document-printlandscape-property-visio%28Office.15%29.aspx)|
|[PrintPagesAcross](http://msdn.microsoft.com/library/document-printpagesacross-property-visio%28Office.15%29.aspx)|
|[PrintPagesDown](http://msdn.microsoft.com/library/document-printpagesdown-property-visio%28Office.15%29.aspx)|
|[PrintScale](http://msdn.microsoft.com/library/document-printscale-property-visio%28Office.15%29.aspx)|
|[ProgID](http://msdn.microsoft.com/library/document-progid-property-visio%28Office.15%29.aspx)|
|[Protection](http://msdn.microsoft.com/library/document-protection-property-visio%28Office.15%29.aspx)|
|[ReadOnly](http://msdn.microsoft.com/library/document-readonly-property-visio%28Office.15%29.aspx)|
|[RemovePersonalInformation](http://msdn.microsoft.com/library/document-removepersonalinformation-property-visio%28Office.15%29.aspx)|
|[RightMargin](http://msdn.microsoft.com/library/document-rightmargin-property-visio%28Office.15%29.aspx)|
|[Saved](http://msdn.microsoft.com/library/document-saved-property-visio%28Office.15%29.aspx)|
|[SavePreviewMode](http://msdn.microsoft.com/library/document-savepreviewmode-property-visio%28Office.15%29.aspx)|
|[ServerPublishOptions](http://msdn.microsoft.com/library/document-serverpublishoptions-property-visio%28Office.15%29.aspx)|
|[SharedWorkspace](http://msdn.microsoft.com/library/document-sharedworkspace-property-visio%28Office.15%29.aspx)|
|[SnapAngles](http://msdn.microsoft.com/library/document-snapangles-property-visio%28Office.15%29.aspx)|
|[SnapEnabled](http://msdn.microsoft.com/library/document-snapenabled-property-visio%28Office.15%29.aspx)|
|[SnapExtensions](http://msdn.microsoft.com/library/document-snapextensions-property-visio%28Office.15%29.aspx)|
|[SnapSettings](http://msdn.microsoft.com/library/document-snapsettings-property-visio%28Office.15%29.aspx)|
|[SolutionXMLElement](http://msdn.microsoft.com/library/document-solutionxmlelement-property-visio%28Office.15%29.aspx)|
|[SolutionXMLElementCount](http://msdn.microsoft.com/library/document-solutionxmlelementcount-property-visio%28Office.15%29.aspx)|
|[SolutionXMLElementExists](http://msdn.microsoft.com/library/document-solutionxmlelementexists-property-visio%28Office.15%29.aspx)|
|[SolutionXMLElementName](http://msdn.microsoft.com/library/document-solutionxmlelementname-property-visio%28Office.15%29.aspx)|
|[Stat](http://msdn.microsoft.com/library/document-stat-property-visio%28Office.15%29.aspx)|
|[Styles](http://msdn.microsoft.com/library/document-styles-property-visio%28Office.15%29.aspx)|
|[Subject](http://msdn.microsoft.com/library/document-subject-property-visio%28Office.15%29.aspx)|
|[Sync](http://msdn.microsoft.com/library/document-sync-property-visio%28Office.15%29.aspx)|
|[Template](http://msdn.microsoft.com/library/document-template-property-visio%28Office.15%29.aspx)|
|[Time](http://msdn.microsoft.com/library/document-time-property-visio%28Office.15%29.aspx)|
|[TimeCreated](http://msdn.microsoft.com/library/document-timecreated-property-visio%28Office.15%29.aspx)|
|[TimeEdited](http://msdn.microsoft.com/library/document-timeedited-property-visio%28Office.15%29.aspx)|
|[TimePrinted](http://msdn.microsoft.com/library/document-timeprinted-property-visio%28Office.15%29.aspx)|
|[TimeSaved](http://msdn.microsoft.com/library/document-timesaved-property-visio%28Office.15%29.aspx)|
|[Title](http://msdn.microsoft.com/library/document-title-property-visio%28Office.15%29.aspx)|
|[TopMargin](http://msdn.microsoft.com/library/document-topmargin-property-visio%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/document-type-property-visio%28Office.15%29.aspx)|
|[UndoEnabled](http://msdn.microsoft.com/library/document-undoenabled-property-visio%28Office.15%29.aspx)|
|[UserCustomUI](http://msdn.microsoft.com/library/document-usercustomui-property-visio%28Office.15%29.aspx)|
|[Validation](http://msdn.microsoft.com/library/document-validation-property-visio%28Office.15%29.aspx)|
|[VBProject](http://msdn.microsoft.com/library/document-vbproject-property-visio%28Office.15%29.aspx)|
|[VBProjectData](http://msdn.microsoft.com/library/document-vbprojectdata-property-visio%28Office.15%29.aspx)|
|[Version](http://msdn.microsoft.com/library/document-version-property-visio%28Office.15%29.aspx)|
|[ZoomBehavior](http://msdn.microsoft.com/library/document-zoombehavior-property-visio%28Office.15%29.aspx)|
|[Permission](http://msdn.microsoft.com/library/document-permission-property-visio%28Office.15%29.aspx)|

