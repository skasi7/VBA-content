---
title: Document Object (Word)
keywords: vbawd10.chm2411
f1_keywords:
- vbawd10.chm2411
ms.prod: WORD
api_name:
- Word.Document
ms.assetid: 8d83487a-2345-a036-a916-971c9db5b7fb
---


# Document Object (Word)

Represents a document. The  **Document** object is a member of the **[Documents](http://msdn.microsoft.com/library/documents-object-word%28Office.15%29.aspx)** collection. The **Documents** collection contains all the **Document** objects that are currently open in Word.


## Remarks

Use  **[Documents](http://msdn.microsoft.com/library/application-documents-property-word%28Office.15%29.aspx)** (Index), where Index is the document name or index number to return a single **Document** object. The following example closes the document named "Report.doc" without saving changes.


```
Documents("Report.doc").Close SaveChanges:=wdDoNotSaveChanges
```

The index number represents the position of the document in the  **Documents** collection. The following example activates the first document in the **Documents** collection.




```
Documents(1).Activate
```

Using ActiveDocument

You can use the  **[ActiveDocument](http://msdn.microsoft.com/library/application-activedocument-property-word%28Office.15%29.aspx)** property to refer to the document with the focus. The following example uses the **[Activate](http://msdn.microsoft.com/library/document-activate-method-word%28Office.15%29.aspx)** method to activate the document named "Document 1." The example also sets the page orientation to landscape mode and then prints the document.




```
Documents("Document1").Activate 
ActiveDocument.PageSetup.Orientation = wdOrientLandscape 
ActiveDocument.PrintOut
```


## Members


### Events



|**Name**|
|:-----|
|[BuildingBlockInsert](http://msdn.microsoft.com/library/document-buildingblockinsert-event-word%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/document-close-event-word%28Office.15%29.aspx)|
|[ContentControlAfterAdd](http://msdn.microsoft.com/library/document-contentcontrolafteradd-event-word%28Office.15%29.aspx)|
|[ContentControlBeforeContentUpdate](http://msdn.microsoft.com/library/document-contentcontrolbeforecontentupdate-event-word%28Office.15%29.aspx)|
|[ContentControlBeforeDelete](http://msdn.microsoft.com/library/document-contentcontrolbeforedelete-event-word%28Office.15%29.aspx)|
|[ContentControlBeforeStoreUpdate](http://msdn.microsoft.com/library/document-contentcontrolbeforestoreupdate-event-word%28Office.15%29.aspx)|
|[ContentControlOnEnter](http://msdn.microsoft.com/library/document-contentcontrolonenter-event-word%28Office.15%29.aspx)|
|[ContentControlOnExit](http://msdn.microsoft.com/library/document-contentcontrolonexit-event-word%28Office.15%29.aspx)|
|[New](http://msdn.microsoft.com/library/document-new-event-word%28Office.15%29.aspx)|
|[Open](http://msdn.microsoft.com/library/document-open-event-word%28Office.15%29.aspx)|
|[Sync](http://msdn.microsoft.com/library/document-sync-event-word%28Office.15%29.aspx)|
|[XMLAfterInsert](http://msdn.microsoft.com/library/document-xmlafterinsert-event-word%28Office.15%29.aspx)|
|[XMLBeforeDelete](http://msdn.microsoft.com/library/document-xmlbeforedelete-event-word%28Office.15%29.aspx)|

### Methods



|**Name**|
|:-----|
|[AcceptAllRevisions](http://msdn.microsoft.com/library/document-acceptallrevisions-method-word%28Office.15%29.aspx)|
|[AcceptAllRevisionsShown](http://msdn.microsoft.com/library/document-acceptallrevisionsshown-method-word%28Office.15%29.aspx)|
|[Activate](http://msdn.microsoft.com/library/document-activate-method-word%28Office.15%29.aspx)|
|[AddToFavorites](http://msdn.microsoft.com/library/document-addtofavorites-method-word%28Office.15%29.aspx)|
|[ApplyDocumentTheme](http://msdn.microsoft.com/library/document-applydocumenttheme-method%28Office.15%29.aspx)|
|[ApplyQuickStyleSet2](http://msdn.microsoft.com/library/document-applyquickstyleset2-method-word%28Office.15%29.aspx)|
|[ApplyTheme](http://msdn.microsoft.com/library/document-applytheme-method-word%28Office.15%29.aspx)|
|[AutoFormat](http://msdn.microsoft.com/library/document-autoformat-method-word%28Office.15%29.aspx)|
|[CanCheckin](http://msdn.microsoft.com/library/document-cancheckin-method-word%28Office.15%29.aspx)|
|[CheckConsistency](http://msdn.microsoft.com/library/document-checkconsistency-method-word%28Office.15%29.aspx)|
|[CheckGrammar](http://msdn.microsoft.com/library/document-checkgrammar-method-word%28Office.15%29.aspx)|
|[CheckIn](http://msdn.microsoft.com/library/document-checkin-method-word%28Office.15%29.aspx)|
|[CheckInWithVersion](http://msdn.microsoft.com/library/document-checkinwithversion-method-word%28Office.15%29.aspx)|
|[CheckSpelling](http://msdn.microsoft.com/library/document-checkspelling-method-word%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/document-close-method-word%28Office.15%29.aspx)|
|[ClosePrintPreview](http://msdn.microsoft.com/library/document-closeprintpreview-method-word%28Office.15%29.aspx)|
|[Compare](http://msdn.microsoft.com/library/document-compare-method-word%28Office.15%29.aspx)|
|[ComputeStatistics](http://msdn.microsoft.com/library/document-computestatistics-method-word%28Office.15%29.aspx)|
|[Convert](http://msdn.microsoft.com/library/document-convert-method-word%28Office.15%29.aspx)|
|[ConvertAutoHyphens](http://msdn.microsoft.com/library/document-convertautohyphens-method-word%28Office.15%29.aspx)|
|[ConvertNumbersToText](http://msdn.microsoft.com/library/document-convertnumberstotext-method-word%28Office.15%29.aspx)|
|[ConvertVietDoc](http://msdn.microsoft.com/library/document-convertvietdoc-method-word%28Office.15%29.aspx)|
|[CopyStylesFromTemplate](http://msdn.microsoft.com/library/document-copystylesfromtemplate-method-word%28Office.15%29.aspx)|
|[CountNumberedItems](http://msdn.microsoft.com/library/document-countnumbereditems-method-word%28Office.15%29.aspx)|
|[CreateLetterContent](http://msdn.microsoft.com/library/document-createlettercontent-method-word%28Office.15%29.aspx)|
|[DataForm](http://msdn.microsoft.com/library/document-dataform-method-word%28Office.15%29.aspx)|
|[DeleteAllComments](http://msdn.microsoft.com/library/document-deleteallcomments-method-word%28Office.15%29.aspx)|
|[DeleteAllCommentsShown](http://msdn.microsoft.com/library/document-deleteallcommentsshown-method-word%28Office.15%29.aspx)|
|[DeleteAllEditableRanges](http://msdn.microsoft.com/library/document-deletealleditableranges-method-word%28Office.15%29.aspx)|
|[DeleteAllInkAnnotations](http://msdn.microsoft.com/library/document-deleteallinkannotations-method-word%28Office.15%29.aspx)|
|[DetectLanguage](http://msdn.microsoft.com/library/document-detectlanguage-method-word%28Office.15%29.aspx)|
|[DowngradeDocument](http://msdn.microsoft.com/library/document-downgradedocument-method-word%28Office.15%29.aspx)|
|[EndReview](http://msdn.microsoft.com/library/document-endreview-method-word%28Office.15%29.aspx)|
|[ExportAsFixedFormat](http://msdn.microsoft.com/library/document-exportasfixedformat-method-word%28Office.15%29.aspx)|
|[FitToPages](http://msdn.microsoft.com/library/document-fittopages-method-word%28Office.15%29.aspx)|
|[FollowHyperlink](http://msdn.microsoft.com/library/document-followhyperlink-method-word%28Office.15%29.aspx)|
|[FreezeLayout](http://msdn.microsoft.com/library/document-freezelayout-method-word%28Office.15%29.aspx)|
|[GetCrossReferenceItems](http://msdn.microsoft.com/library/document-getcrossreferenceitems-method-word%28Office.15%29.aspx)|
|[GetLetterContent](http://msdn.microsoft.com/library/document-getlettercontent-method-word%28Office.15%29.aspx)|
|[GetWorkflowTasks](http://msdn.microsoft.com/library/document-getworkflowtasks-method-word%28Office.15%29.aspx)|
|[GetWorkflowTemplates](http://msdn.microsoft.com/library/document-getworkflowtemplates-method-word%28Office.15%29.aspx)|
|[GoTo](http://msdn.microsoft.com/library/document-goto-method-word%28Office.15%29.aspx)|
|[LockServerFile](http://msdn.microsoft.com/library/document-lockserverfile-method-word%28Office.15%29.aspx)|
|[MakeCompatibilityDefault](http://msdn.microsoft.com/library/document-makecompatibilitydefault-method-word%28Office.15%29.aspx)|
|[ManualHyphenation](http://msdn.microsoft.com/library/document-manualhyphenation-method-word%28Office.15%29.aspx)|
|[Merge](http://msdn.microsoft.com/library/document-merge-method-word%28Office.15%29.aspx)|
|[Post](http://msdn.microsoft.com/library/document-post-method-word%28Office.15%29.aspx)|
|[PresentIt](http://msdn.microsoft.com/library/document-presentit-method-word%28Office.15%29.aspx)|
|[PrintOut](http://msdn.microsoft.com/library/document-printout-method-word%28Office.15%29.aspx)|
|[PrintPreview](http://msdn.microsoft.com/library/document-printpreview-method-word%28Office.15%29.aspx)|
|[Protect](http://msdn.microsoft.com/library/document-protect-method-word%28Office.15%29.aspx)|
|[Range](http://msdn.microsoft.com/library/document-range-method-word%28Office.15%29.aspx)|
|[Redo](http://msdn.microsoft.com/library/document-redo-method-word%28Office.15%29.aspx)|
|[RejectAllRevisions](http://msdn.microsoft.com/library/document-rejectallrevisions-method-word%28Office.15%29.aspx)|
|[RejectAllRevisionsShown](http://msdn.microsoft.com/library/document-rejectallrevisionsshown-method-word%28Office.15%29.aspx)|
|[Reload](http://msdn.microsoft.com/library/document-reload-method-word%28Office.15%29.aspx)|
|[ReloadAs](http://msdn.microsoft.com/library/document-reloadas-method-word%28Office.15%29.aspx)|
|[RemoveDocumentInformation](http://msdn.microsoft.com/library/document-removedocumentinformation-method-word%28Office.15%29.aspx)|
|[RemoveLockedStyles](http://msdn.microsoft.com/library/document-removelockedstyles-method-word%28Office.15%29.aspx)|
|[RemoveNumbers](http://msdn.microsoft.com/library/document-removenumbers-method-word%28Office.15%29.aspx)|
|[RemoveTheme](http://msdn.microsoft.com/library/document-removetheme-method-word%28Office.15%29.aspx)|
|[Repaginate](http://msdn.microsoft.com/library/document-repaginate-method-word%28Office.15%29.aspx)|
|[ReplyWithChanges](http://msdn.microsoft.com/library/document-replywithchanges-method-word%28Office.15%29.aspx)|
|[ResetFormFields](http://msdn.microsoft.com/library/document-resetformfields-method-word%28Office.15%29.aspx)|
|[ReturnToLastReadPosition](http://msdn.microsoft.com/library/document-returntolastreadposition-method-word%28Office.15%29.aspx)|
|[RunAutoMacro](http://msdn.microsoft.com/library/document-runautomacro-method-word%28Office.15%29.aspx)|
|[RunLetterWizard](http://msdn.microsoft.com/library/document-runletterwizard-method-word%28Office.15%29.aspx)|
|[Save](http://msdn.microsoft.com/library/document-save-method-word%28Office.15%29.aspx)|
|[SaveAs2](http://msdn.microsoft.com/library/document-saveas2-method-word%28Office.15%29.aspx)|
|[SaveAsQuickStyleSet](http://msdn.microsoft.com/library/document-saveasquickstyleset-method-word%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/document-select-method-word%28Office.15%29.aspx)|
|[SelectAllEditableRanges](http://msdn.microsoft.com/library/document-selectalleditableranges-method-word%28Office.15%29.aspx)|
|[SelectContentControlsByTag](http://msdn.microsoft.com/library/document-selectcontentcontrolsbytag-method-word%28Office.15%29.aspx)|
|[SelectContentControlsByTitle](http://msdn.microsoft.com/library/document-selectcontentcontrolsbytitle-method-word%28Office.15%29.aspx)|
|[SelectLinkedControls](http://msdn.microsoft.com/library/document-selectlinkedcontrols-method-word%28Office.15%29.aspx)|
|[SelectNodes](http://msdn.microsoft.com/library/document-selectnodes-method-word%28Office.15%29.aspx)|
|[SelectSingleNode](http://msdn.microsoft.com/library/document-selectsinglenode-method-word%28Office.15%29.aspx)|
|[SelectUnlinkedControls](http://msdn.microsoft.com/library/document-selectunlinkedcontrols-method-word%28Office.15%29.aspx)|
|[SendFax](http://msdn.microsoft.com/library/document-sendfax-method-word%28Office.15%29.aspx)|
|[SendFaxOverInternet](http://msdn.microsoft.com/library/document-sendfaxoverinternet-method-word%28Office.15%29.aspx)|
|[SendForReview](http://msdn.microsoft.com/library/document-sendforreview-method-word%28Office.15%29.aspx)|
|[SendMail](http://msdn.microsoft.com/library/document-sendmail-method-word%28Office.15%29.aspx)|
|[SetCompatibilityMode](http://msdn.microsoft.com/library/document-setcompatibilitymode-method-word%28Office.15%29.aspx)|
|[SetDefaultTableStyle](http://msdn.microsoft.com/library/document-setdefaulttablestyle-method-word%28Office.15%29.aspx)|
|[SetLetterContent](http://msdn.microsoft.com/library/document-setlettercontent-method-word%28Office.15%29.aspx)|
|[SetPasswordEncryptionOptions](http://msdn.microsoft.com/library/document-setpasswordencryptionoptions-method-word%28Office.15%29.aspx)|
|[ToggleFormsDesign](http://msdn.microsoft.com/library/document-toggleformsdesign-method-word%28Office.15%29.aspx)|
|[TransformDocument](http://msdn.microsoft.com/library/document-transformdocument-method-word%28Office.15%29.aspx)|
|[Undo](http://msdn.microsoft.com/library/document-undo-method-word%28Office.15%29.aspx)|
|[UndoClear](http://msdn.microsoft.com/library/document-undoclear-method-word%28Office.15%29.aspx)|
|[Unprotect](http://msdn.microsoft.com/library/document-unprotect-method-word%28Office.15%29.aspx)|
|[UpdateStyles](http://msdn.microsoft.com/library/document-updatestyles-method-word%28Office.15%29.aspx)|
|[ViewCode](http://msdn.microsoft.com/library/document-viewcode-method-word%28Office.15%29.aspx)|
|[ViewPropertyBrowser](http://msdn.microsoft.com/library/document-viewpropertybrowser-method-word%28Office.15%29.aspx)|
|[WebPagePreview](http://msdn.microsoft.com/library/document-webpagepreview-method-word%28Office.15%29.aspx)|

### Properties



|**Name**|
|:-----|
|[ActiveTheme](http://msdn.microsoft.com/library/document-activetheme-property-word%28Office.15%29.aspx)|
|[ActiveThemeDisplayName](http://msdn.microsoft.com/library/document-activethemedisplayname-property-word%28Office.15%29.aspx)|
|[ActiveWindow](http://msdn.microsoft.com/library/document-activewindow-property-word%28Office.15%29.aspx)|
|[ActiveWritingStyle](http://msdn.microsoft.com/library/document-activewritingstyle-property-word%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/document-application-property-word%28Office.15%29.aspx)|
|[AttachedTemplate](http://msdn.microsoft.com/library/document-attachedtemplate-property-word%28Office.15%29.aspx)|
|[AutoFormatOverride](http://msdn.microsoft.com/library/document-autoformatoverride-property-word%28Office.15%29.aspx)|
|[AutoHyphenation](http://msdn.microsoft.com/library/document-autohyphenation-property-word%28Office.15%29.aspx)|
|[Background](http://msdn.microsoft.com/library/document-background-property-word%28Office.15%29.aspx)|
|[Bibliography](http://msdn.microsoft.com/library/document-bibliography-property-word%28Office.15%29.aspx)|
|[Bookmarks](http://msdn.microsoft.com/library/document-bookmarks-property-word%28Office.15%29.aspx)|
|[Broadcast](http://msdn.microsoft.com/library/document-broadcast-property-word%28Office.15%29.aspx)|
|[BuiltInDocumentProperties](http://msdn.microsoft.com/library/document-builtindocumentproperties-property-word%28Office.15%29.aspx)|
|[Characters](http://msdn.microsoft.com/library/document-characters-property-word%28Office.15%29.aspx)|
|[ChartDataPointTrack](http://msdn.microsoft.com/library/document-chartdatapointtrack-property-word%28Office.15%29.aspx)|
|[ClickAndTypeParagraphStyle](http://msdn.microsoft.com/library/document-clickandtypeparagraphstyle-property-word%28Office.15%29.aspx)|
|[CoAuthoring](http://msdn.microsoft.com/library/document-coauthoring-property-word%28Office.15%29.aspx)|
|[CodeName](http://msdn.microsoft.com/library/document-codename-property-word%28Office.15%29.aspx)|
|[CommandBars](http://msdn.microsoft.com/library/document-commandbars-property-word%28Office.15%29.aspx)|
|[Comments](http://msdn.microsoft.com/library/document-comments-property-word%28Office.15%29.aspx)|
|[Compatibility](http://msdn.microsoft.com/library/document-compatibility-property-word%28Office.15%29.aspx)|
|[CompatibilityMode](http://msdn.microsoft.com/library/document-compatibilitymode-property-word%28Office.15%29.aspx)|
|[ConsecutiveHyphensLimit](http://msdn.microsoft.com/library/document-consecutivehyphenslimit-property-word%28Office.15%29.aspx)|
|[Container](http://msdn.microsoft.com/library/document-container-property-word%28Office.15%29.aspx)|
|[Content](http://msdn.microsoft.com/library/document-content-property-word%28Office.15%29.aspx)|
|[ContentControls](http://msdn.microsoft.com/library/document-contentcontrols-property-word%28Office.15%29.aspx)|
|[ContentTypeProperties](http://msdn.microsoft.com/library/document-contenttypeproperties-property-word%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/document-creator-property-word%28Office.15%29.aspx)|
|[CurrentRsid](http://msdn.microsoft.com/library/document-currentrsid-property-word%28Office.15%29.aspx)|
|[CustomDocumentProperties](http://msdn.microsoft.com/library/document-customdocumentproperties-property-word%28Office.15%29.aspx)|
|[CustomXMLParts](http://msdn.microsoft.com/library/document-customxmlparts-property-word%28Office.15%29.aspx)|
|[DefaultTableStyle](http://msdn.microsoft.com/library/document-defaulttablestyle-property-word%28Office.15%29.aspx)|
|[DefaultTabStop](http://msdn.microsoft.com/library/document-defaulttabstop-property-word%28Office.15%29.aspx)|
|[DefaultTargetFrame](http://msdn.microsoft.com/library/document-defaulttargetframe-property-word%28Office.15%29.aspx)|
|[DisableFeatures](http://msdn.microsoft.com/library/document-disablefeatures-property-word%28Office.15%29.aspx)|
|[DisableFeaturesIntroducedAfter](http://msdn.microsoft.com/library/document-disablefeaturesintroducedafter-property-word%28Office.15%29.aspx)|
|[DocumentInspectors](http://msdn.microsoft.com/library/document-documentinspectors-property-word%28Office.15%29.aspx)|
|[DocumentLibraryVersions](http://msdn.microsoft.com/library/document-documentlibraryversions-property-word%28Office.15%29.aspx)|
|[DocumentTheme](http://msdn.microsoft.com/library/document-documenttheme-property-word%28Office.15%29.aspx)|
|[DoNotEmbedSystemFonts](http://msdn.microsoft.com/library/document-donotembedsystemfonts-property-word%28Office.15%29.aspx)|
|[Email](http://msdn.microsoft.com/library/document-email-property-word%28Office.15%29.aspx)|
|[EmbedLinguisticData](http://msdn.microsoft.com/library/document-embedlinguisticdata-property-word%28Office.15%29.aspx)|
|[EmbedTrueTypeFonts](http://msdn.microsoft.com/library/document-embedtruetypefonts-property-word%28Office.15%29.aspx)|
|[EncryptionProvider](http://msdn.microsoft.com/library/document-encryptionprovider-property-word%28Office.15%29.aspx)|
|[Endnotes](http://msdn.microsoft.com/library/document-endnotes-property-word%28Office.15%29.aspx)|
|[EnforceStyle](http://msdn.microsoft.com/library/document-enforcestyle-property-word%28Office.15%29.aspx)|
|[Envelope](http://msdn.microsoft.com/library/document-envelope-property-word%28Office.15%29.aspx)|
|[FarEastLineBreakLanguage](http://msdn.microsoft.com/library/document-fareastlinebreaklanguage-property-word%28Office.15%29.aspx)|
|[FarEastLineBreakLevel](http://msdn.microsoft.com/library/document-fareastlinebreaklevel-property-word%28Office.15%29.aspx)|
|[Fields](http://msdn.microsoft.com/library/document-fields-property-word%28Office.15%29.aspx)|
|[Final](http://msdn.microsoft.com/library/document-final-property-word%28Office.15%29.aspx)|
|[Footnotes](http://msdn.microsoft.com/library/document-footnotes-property-word%28Office.15%29.aspx)|
|[FormattingShowClear](http://msdn.microsoft.com/library/document-formattingshowclear-property-word%28Office.15%29.aspx)|
|[FormattingShowFilter](http://msdn.microsoft.com/library/document-formattingshowfilter-property-word%28Office.15%29.aspx)|
|[FormattingShowFont](http://msdn.microsoft.com/library/document-formattingshowfont-property-word%28Office.15%29.aspx)|
|[FormattingShowNextLevel](http://msdn.microsoft.com/library/document-formattingshownextlevel-property-word%28Office.15%29.aspx)|
|[FormattingShowNumbering](http://msdn.microsoft.com/library/document-formattingshownumbering-property-word%28Office.15%29.aspx)|
|[FormattingShowParagraph](http://msdn.microsoft.com/library/document-formattingshowparagraph-property-word%28Office.15%29.aspx)|
|[FormattingShowUserStyleName](http://msdn.microsoft.com/library/document-formattingshowuserstylename-property-word%28Office.15%29.aspx)|
|[FormFields](http://msdn.microsoft.com/library/document-formfields-property-word%28Office.15%29.aspx)|
|[FormsDesign](http://msdn.microsoft.com/library/document-formsdesign-property-word%28Office.15%29.aspx)|
|[Frames](http://msdn.microsoft.com/library/document-frames-property-word%28Office.15%29.aspx)|
|[Frameset](http://msdn.microsoft.com/library/document-frameset-property-word%28Office.15%29.aspx)|
|[FullName](http://msdn.microsoft.com/library/document-fullname-property-word%28Office.15%29.aspx)|
|[GrammarChecked](http://msdn.microsoft.com/library/document-grammarchecked-property-word%28Office.15%29.aspx)|
|[GrammaticalErrors](http://msdn.microsoft.com/library/document-grammaticalerrors-property-word%28Office.15%29.aspx)|
|[GridDistanceHorizontal](http://msdn.microsoft.com/library/document-griddistancehorizontal-property-word%28Office.15%29.aspx)|
|[GridDistanceVertical](http://msdn.microsoft.com/library/document-griddistancevertical-property-word%28Office.15%29.aspx)|
|[GridOriginFromMargin](http://msdn.microsoft.com/library/document-gridoriginfrommargin-property-word%28Office.15%29.aspx)|
|[GridOriginHorizontal](http://msdn.microsoft.com/library/document-gridoriginhorizontal-property-word%28Office.15%29.aspx)|
|[GridOriginVertical](http://msdn.microsoft.com/library/document-gridoriginvertical-property-word%28Office.15%29.aspx)|
|[GridSpaceBetweenHorizontalLines](http://msdn.microsoft.com/library/document-gridspacebetweenhorizontallines-property-word%28Office.15%29.aspx)|
|[GridSpaceBetweenVerticalLines](http://msdn.microsoft.com/library/document-gridspacebetweenverticallines-property-word%28Office.15%29.aspx)|
|[HasPassword](http://msdn.microsoft.com/library/document-haspassword-property-word%28Office.15%29.aspx)|
|[HasVBProject](http://msdn.microsoft.com/library/document-hasvbproject-property-word%28Office.15%29.aspx)|
|[HTMLDivisions](http://msdn.microsoft.com/library/document-htmldivisions-property-word%28Office.15%29.aspx)|
|[Hyperlinks](http://msdn.microsoft.com/library/document-hyperlinks-property-word%28Office.15%29.aspx)|
|[HyphenateCaps](http://msdn.microsoft.com/library/document-hyphenatecaps-property-word%28Office.15%29.aspx)|
|[HyphenationZone](http://msdn.microsoft.com/library/document-hyphenationzone-property-word%28Office.15%29.aspx)|
|[Indexes](http://msdn.microsoft.com/library/document-indexes-property-word%28Office.15%29.aspx)|
|[InlineShapes](http://msdn.microsoft.com/library/document-inlineshapes-property-word%28Office.15%29.aspx)|
|[IsInAutosave](http://msdn.microsoft.com/library/document-isinautosave-property-word%28Office.15%29.aspx)|
|[IsMasterDocument](http://msdn.microsoft.com/library/document-ismasterdocument-property-word%28Office.15%29.aspx)|
|[IsSubdocument](http://msdn.microsoft.com/library/document-issubdocument-property-word%28Office.15%29.aspx)|
|[JustificationMode](http://msdn.microsoft.com/library/document-justificationmode-property-word%28Office.15%29.aspx)|
|[KerningByAlgorithm](http://msdn.microsoft.com/library/document-kerningbyalgorithm-property-word%28Office.15%29.aspx)|
|[Kind](http://msdn.microsoft.com/library/document-kind-property-word%28Office.15%29.aspx)|
|[LanguageDetected](http://msdn.microsoft.com/library/document-languagedetected-property-word%28Office.15%29.aspx)|
|[ListParagraphs](http://msdn.microsoft.com/library/document-listparagraphs-property-word%28Office.15%29.aspx)|
|[Lists](http://msdn.microsoft.com/library/document-lists-property-word%28Office.15%29.aspx)|
|[ListTemplates](http://msdn.microsoft.com/library/document-listtemplates-property-word%28Office.15%29.aspx)|
|[LockQuickStyleSet](http://msdn.microsoft.com/library/document-lockquickstyleset-property-word%28Office.15%29.aspx)|
|[LockTheme](http://msdn.microsoft.com/library/document-locktheme-property-word%28Office.15%29.aspx)|
|[MailEnvelope](http://msdn.microsoft.com/library/document-mailenvelope-property-word%28Office.15%29.aspx)|
|[MailMerge](http://msdn.microsoft.com/library/document-mailmerge-property-word%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/document-name-property-word%28Office.15%29.aspx)|
|[NoLineBreakAfter](http://msdn.microsoft.com/library/document-nolinebreakafter-property-word%28Office.15%29.aspx)|
|[NoLineBreakBefore](http://msdn.microsoft.com/library/document-nolinebreakbefore-property-word%28Office.15%29.aspx)|
|[OMathBreakBin](http://msdn.microsoft.com/library/document-omathbreakbin-property-word%28Office.15%29.aspx)|
|[OMathBreakSub](http://msdn.microsoft.com/library/document-omathbreaksub-property-word%28Office.15%29.aspx)|
|[OMathFontName](http://msdn.microsoft.com/library/document-omathfontname-property-word%28Office.15%29.aspx)|
|[OMathIntSubSupLim](http://msdn.microsoft.com/library/document-omathintsubsuplim-property-word%28Office.15%29.aspx)|
|[OMathJc](http://msdn.microsoft.com/library/document-omathjc-property-word%28Office.15%29.aspx)|
|[OMathLeftMargin](http://msdn.microsoft.com/library/document-omathleftmargin-property-word%28Office.15%29.aspx)|
|[OMathNarySupSubLim](http://msdn.microsoft.com/library/document-omathnarysupsublim-property-word%28Office.15%29.aspx)|
|[OMathRightMargin](http://msdn.microsoft.com/library/document-omathrightmargin-property-word%28Office.15%29.aspx)|
|[OMaths](http://msdn.microsoft.com/library/document-omaths-property-word%28Office.15%29.aspx)|
|[OMathSmallFrac](http://msdn.microsoft.com/library/document-omathsmallfrac-property-word%28Office.15%29.aspx)|
|[OMathWrap](http://msdn.microsoft.com/library/document-omathwrap-property-word%28Office.15%29.aspx)|
|[OpenEncoding](http://msdn.microsoft.com/library/document-openencoding-property-word%28Office.15%29.aspx)|
|[OptimizeForWord97](http://msdn.microsoft.com/library/document-optimizeforword97-property-word%28Office.15%29.aspx)|
|[OriginalDocumentTitle](http://msdn.microsoft.com/library/document-originaldocumenttitle-property-word%28Office.15%29.aspx)|
|[PageSetup](http://msdn.microsoft.com/library/document-pagesetup-property-word%28Office.15%29.aspx)|
|[Paragraphs](http://msdn.microsoft.com/library/document-paragraphs-property-word%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/document-parent-property-word%28Office.15%29.aspx)|
|[Password](http://msdn.microsoft.com/library/document-password-property-word%28Office.15%29.aspx)|
|[PasswordEncryptionAlgorithm](http://msdn.microsoft.com/library/document-passwordencryptionalgorithm-property-word%28Office.15%29.aspx)|
|[PasswordEncryptionFileProperties](http://msdn.microsoft.com/library/document-passwordencryptionfileproperties-property-word%28Office.15%29.aspx)|
|[PasswordEncryptionKeyLength](http://msdn.microsoft.com/library/document-passwordencryptionkeylength-property-word%28Office.15%29.aspx)|
|[PasswordEncryptionProvider](http://msdn.microsoft.com/library/document-passwordencryptionprovider-property-word%28Office.15%29.aspx)|
|[Path](http://msdn.microsoft.com/library/document-path-property-word%28Office.15%29.aspx)|
|[Permission](http://msdn.microsoft.com/library/document-permission-property-word%28Office.15%29.aspx)|
|[PrintFormsData](http://msdn.microsoft.com/library/document-printformsdata-property-word%28Office.15%29.aspx)|
|[PrintPostScriptOverText](http://msdn.microsoft.com/library/document-printpostscriptovertext-property-word%28Office.15%29.aspx)|
|[PrintRevisions](http://msdn.microsoft.com/library/document-printrevisions-property-word%28Office.15%29.aspx)|
|[ProtectionType](http://msdn.microsoft.com/library/document-protectiontype-property-word%28Office.15%29.aspx)|
|[ReadabilityStatistics](http://msdn.microsoft.com/library/document-readabilitystatistics-property-word%28Office.15%29.aspx)|
|[ReadingLayoutSizeX](http://msdn.microsoft.com/library/document-readinglayoutsizex-property-word%28Office.15%29.aspx)|
|[ReadingLayoutSizeY](http://msdn.microsoft.com/library/document-readinglayoutsizey-property-word%28Office.15%29.aspx)|
|[ReadingModeLayoutFrozen](http://msdn.microsoft.com/library/document-readingmodelayoutfrozen-property-word%28Office.15%29.aspx)|
|[ReadOnly](http://msdn.microsoft.com/library/document-readonly-property-word%28Office.15%29.aspx)|
|[ReadOnlyRecommended](http://msdn.microsoft.com/library/document-readonlyrecommended-property-word%28Office.15%29.aspx)|
|[RemoveDateAndTime](http://msdn.microsoft.com/library/document-removedateandtime-property-word%28Office.15%29.aspx)|
|[RemovePersonalInformation](http://msdn.microsoft.com/library/document-removepersonalinformation-property-word%28Office.15%29.aspx)|
|[Research](http://msdn.microsoft.com/library/document-research-property-word%28Office.15%29.aspx)|
|[RevisedDocumentTitle](http://msdn.microsoft.com/library/document-reviseddocumenttitle-property-word%28Office.15%29.aspx)|
|[Revisions](http://msdn.microsoft.com/library/document-revisions-property-word%28Office.15%29.aspx)|
|[Saved](http://msdn.microsoft.com/library/document-saved-property-word%28Office.15%29.aspx)|
|[SaveEncoding](http://msdn.microsoft.com/library/document-saveencoding-property-word%28Office.15%29.aspx)|
|[SaveFormat](http://msdn.microsoft.com/library/document-saveformat-property-word%28Office.15%29.aspx)|
|[SaveFormsData](http://msdn.microsoft.com/library/document-saveformsdata-property-word%28Office.15%29.aspx)|
|[SaveSubsetFonts](http://msdn.microsoft.com/library/document-savesubsetfonts-property-word%28Office.15%29.aspx)|
|[Scripts](http://msdn.microsoft.com/library/document-scripts-property-word%28Office.15%29.aspx)|
|[Sections](http://msdn.microsoft.com/library/document-sections-property-word%28Office.15%29.aspx)|
|[Sentences](http://msdn.microsoft.com/library/document-sentences-property-word%28Office.15%29.aspx)|
|[ServerPolicy](http://msdn.microsoft.com/library/document-serverpolicy-property-word%28Office.15%29.aspx)|
|[Shapes](http://msdn.microsoft.com/library/document-shapes-property-word%28Office.15%29.aspx)|
|[ShowGrammaticalErrors](http://msdn.microsoft.com/library/document-showgrammaticalerrors-property-word%28Office.15%29.aspx)|
|[ShowSpellingErrors](http://msdn.microsoft.com/library/document-showspellingerrors-property-word%28Office.15%29.aspx)|
|[Signatures](http://msdn.microsoft.com/library/document-signatures-property-word%28Office.15%29.aspx)|
|[SmartDocument](http://msdn.microsoft.com/library/document-smartdocument-property-word%28Office.15%29.aspx)|
|[SnapToGrid](http://msdn.microsoft.com/library/document-snaptogrid-property-word%28Office.15%29.aspx)|
|[SnapToShapes](http://msdn.microsoft.com/library/document-snaptoshapes-property-word%28Office.15%29.aspx)|
|[SpellingChecked](http://msdn.microsoft.com/library/document-spellingchecked-property-word%28Office.15%29.aspx)|
|[SpellingErrors](http://msdn.microsoft.com/library/document-spellingerrors-property-word%28Office.15%29.aspx)|
|[StoryRanges](http://msdn.microsoft.com/library/document-storyranges-property-word%28Office.15%29.aspx)|
|[Styles](http://msdn.microsoft.com/library/document-styles-property-word%28Office.15%29.aspx)|
|[StyleSheets](http://msdn.microsoft.com/library/document-stylesheets-property-word%28Office.15%29.aspx)|
|[StyleSortMethod](http://msdn.microsoft.com/library/document-stylesortmethod-property-word%28Office.15%29.aspx)|
|[Subdocuments](http://msdn.microsoft.com/library/document-subdocuments-property-word%28Office.15%29.aspx)|
|[Sync](http://msdn.microsoft.com/library/document-sync-property-word%28Office.15%29.aspx)|
|[Tables](http://msdn.microsoft.com/library/document-tables-property-word%28Office.15%29.aspx)|
|[TablesOfAuthorities](http://msdn.microsoft.com/library/document-tablesofauthorities-property-word%28Office.15%29.aspx)|
|[TablesOfAuthoritiesCategories](http://msdn.microsoft.com/library/document-tablesofauthoritiescategories-property-word%28Office.15%29.aspx)|
|[TablesOfContents](http://msdn.microsoft.com/library/document-tablesofcontents-property-word%28Office.15%29.aspx)|
|[TablesOfFigures](http://msdn.microsoft.com/library/document-tablesoffigures-property-word%28Office.15%29.aspx)|
|[TextEncoding](http://msdn.microsoft.com/library/document-textencoding-property-word%28Office.15%29.aspx)|
|[TextLineEnding](http://msdn.microsoft.com/library/document-textlineending-property-word%28Office.15%29.aspx)|
|[TrackFormatting](http://msdn.microsoft.com/library/document-trackformatting-property-word%28Office.15%29.aspx)|
|[TrackMoves](http://msdn.microsoft.com/library/document-trackmoves-property-word%28Office.15%29.aspx)|
|[TrackRevisions](http://msdn.microsoft.com/library/document-trackrevisions-property-word%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/document-type-property-word%28Office.15%29.aspx)|
|[UpdateStylesOnOpen](http://msdn.microsoft.com/library/document-updatestylesonopen-property-word%28Office.15%29.aspx)|
|[UseMathDefaults](http://msdn.microsoft.com/library/document-usemathdefaults-property-word%28Office.15%29.aspx)|
|[UserControl](http://msdn.microsoft.com/library/document-usercontrol-property-word%28Office.15%29.aspx)|
|[Variables](http://msdn.microsoft.com/library/document-variables-property-word%28Office.15%29.aspx)|
|[VBASigned](http://msdn.microsoft.com/library/document-vbasigned-property-word%28Office.15%29.aspx)|
|[VBProject](http://msdn.microsoft.com/library/document-vbproject-property-word%28Office.15%29.aspx)|
|[WebOptions](http://msdn.microsoft.com/library/document-weboptions-property-word%28Office.15%29.aspx)|
|[Windows](http://msdn.microsoft.com/library/document-windows-property-word%28Office.15%29.aspx)|
|[WordOpenXML](http://msdn.microsoft.com/library/document-wordopenxml-property-word%28Office.15%29.aspx)|
|[Words](http://msdn.microsoft.com/library/document-words-property-word%28Office.15%29.aspx)|
|[WritePassword](http://msdn.microsoft.com/library/document-writepassword-property-word%28Office.15%29.aspx)|
|[WriteReserved](http://msdn.microsoft.com/library/document-writereserved-property-word%28Office.15%29.aspx)|
|[XMLSaveThroughXSLT](http://msdn.microsoft.com/library/document-xmlsavethroughxslt-property-word%28Office.15%29.aspx)|
|[XMLSchemaReferences](http://msdn.microsoft.com/library/document-xmlschemareferences-property-word%28Office.15%29.aspx)|
|[XMLShowAdvancedErrors](http://msdn.microsoft.com/library/document-xmlshowadvancederrors-property-word%28Office.15%29.aspx)|
|[XMLUseXSLTWhenSaving](http://msdn.microsoft.com/library/document-xmlusexsltwhensaving-property-word%28Office.15%29.aspx)|

## See also


#### Other resources


<<<<<<< HEAD
[Word Object Model Reference](http://msdn.microsoft.com/library/object-model-word-vba-reference%28Office.15%29.aspx)
=======
[Word Object Model Reference](http://msdn.microsoft.com/library/object-model-word-vba-reference%28Office.15%29.aspx)

>>>>>>> d7667e83d23dbf8ebf5bf068ba6fed14c840c0f5

