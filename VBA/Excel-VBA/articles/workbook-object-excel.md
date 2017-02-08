---
title: Workbook Object (Excel)
keywords: vbaxl10.chm198072
f1_keywords:
- vbaxl10.chm198072
ms.prod: EXCEL
api_name:
- Excel.Workbook
ms.assetid: 8c00aa60-c974-eed3-0812-3c9625eb0d4c
---


# Workbook Object (Excel)

Represents a Microsoft Excel workbook.


## Remarks

The  **Workbook** object is a member of the [Workbooks](http://msdn.microsoft.com/library/workbooks-object-excel%28Office.15%29.aspx) collection. The **Workbooks** collection contains all the **Workbook** objects currently open in Microsoft Excel.


### ThisWorkbook Property

The  [ThisWorkbook](http://msdn.microsoft.com/library/application-thisworkbook-property-excel%28Office.15%29.aspx) property returns the workbook where the Visual Basic code is running. In most cases, this is the same as the active workbook. However, if the Visual Basic code is part of an add-in, the **ThisWorkbook** property won't return the active workbook. In this case, the active workbook is the workbook calling the add-in, whereas the **ThisWorkbook** property returns the add-in workbook.

If you'll be creating an add-in from your Visual Basic code, you should use the  **ThisWorkbook** property to qualify any statement that must be run on the workbook you compile into the add-in.


## Example

Use  **Workbooks** ( _index_ ), where _index_ is the workbook name or index number, to return a single [Workbook](workbook-object-excel.md) object. The following example activates workbook one.


```
Workbooks(1).Activate
```

The index number denotes the order in which the workbooks were opened or created.  `Workbooks(1)` is the first workbook created, and `Workbooks(Workbooks.Count)` is the last one created. Activating a workbook doesn't change its index number. All workbooks are included in the index count, even if they're hidden.



The  **[Name](http://msdn.microsoft.com/library/workbook-name-property-excel%28Office.15%29.aspx)** property returns the workbook name. You cannot set the name by using this property; if you need to change the name, use the **[SaveAs](http://msdn.microsoft.com/library/workbook-saveas-method-excel%28Office.15%29.aspx)** method to save the workbook under a different name. The following example activates Sheet1 in the workbook named Cogs.xls (the workbook must already be open in Microsoft Excel).




```
Workbooks("Cogs.xls").Worksheets("Sheet1").Activate
```

The  **[ActiveWorkbook](http://msdn.microsoft.com/library/application-activeworkbook-property-excel%28Office.15%29.aspx)** property returns the workbook that's currently active. The following example sets the name of the author for the active workbook.






```
ActiveWorkbook.Author = "Jean Selva"
```

 **Sample code provided by:** Holy Macro! Books, [Holy Macro! It's 2,500 Excel VBA Examples](http://www.mrexcel.com/store/index.php?l=product_detail&amp;p=1)

This example emails a worksheet tab from the active workbook using a specified email address and subject. To run this code, the active worksheet must contain the email address in cell A1, the subject in cell B1, and the name of the worksheet to send in cell C1.




```
Sub SendTab()
   'Declare and initialize your variables, and turn off screen updating.
   Dim wks As Worksheet
   Application.ScreenUpdating = False
   Set wks = ActiveSheet
   
   'Copy the target worksheet, specified in cell C1, to the clipboard.
   Worksheets(Range("C1").Value).Copy
   
   'Send the content in the clipboard to the email account specified in cell A1,
   'using the subject line specified in cell B1.
   ActiveWorkbook.SendMail wks.Range("A1").Value, wks.Range("B1").Value
   
   'Do not save changes and turn screen updating back on.
   ActiveWorkbook.Close savechanges:=False
   Application.ScreenUpdating = True
End Sub
```


## Events



|**Name**|
|:-----|
|[Activate](http://msdn.microsoft.com/library/workbook-activate-event-excel%28Office.15%29.aspx)|
|[AddinInstall](http://msdn.microsoft.com/library/workbook-addininstall-event-excel%28Office.15%29.aspx)|
|[AddinUninstall](http://msdn.microsoft.com/library/workbook-addinuninstall-event-excel%28Office.15%29.aspx)|
|[AfterSave](http://msdn.microsoft.com/library/workbook-aftersave-event-excel%28Office.15%29.aspx)|
|[AfterXmlExport](http://msdn.microsoft.com/library/workbook-afterxmlexport-event-excel%28Office.15%29.aspx)|
|[AfterXmlImport](http://msdn.microsoft.com/library/workbook-afterxmlimport-event-excel%28Office.15%29.aspx)|
|[BeforeClose](http://msdn.microsoft.com/library/workbook-beforeclose-event-excel%28Office.15%29.aspx)|
|[BeforePrint](http://msdn.microsoft.com/library/workbook-beforeprint-event-excel%28Office.15%29.aspx)|
|[BeforeSave](http://msdn.microsoft.com/library/workbook-beforesave-event-excel%28Office.15%29.aspx)|
|[BeforeXmlExport](http://msdn.microsoft.com/library/workbook-beforexmlexport-event-excel%28Office.15%29.aspx)|
|[BeforeXmlImport](http://msdn.microsoft.com/library/workbook-beforexmlimport-event-excel%28Office.15%29.aspx)|
|[Deactivate](http://msdn.microsoft.com/library/workbook-deactivate-event-excel%28Office.15%29.aspx)|
|[ModelChange](http://msdn.microsoft.com/library/workbook-modelchange-event-excel%28Office.15%29.aspx)|
|[NewChart](http://msdn.microsoft.com/library/workbook-newchart-event-excel%28Office.15%29.aspx)|
|[NewSheet](http://msdn.microsoft.com/library/workbook-newsheet-event-excel%28Office.15%29.aspx)|
|[Open](http://msdn.microsoft.com/library/workbook-open-event-excel%28Office.15%29.aspx)|
|[PivotTableCloseConnection](http://msdn.microsoft.com/library/workbook-pivottablecloseconnection-event-excel%28Office.15%29.aspx)|
|[PivotTableOpenConnection](http://msdn.microsoft.com/library/workbook-pivottableopenconnection-event-excel%28Office.15%29.aspx)|
|[RowsetComplete](http://msdn.microsoft.com/library/workbook-rowsetcomplete-event-excel%28Office.15%29.aspx)|
|[SheetActivate](http://msdn.microsoft.com/library/workbook-sheetactivate-event-excel%28Office.15%29.aspx)|
|[SheetBeforeDelete](http://msdn.microsoft.com/library/workbook-sheetbeforedelete-event-excel%28Office.15%29.aspx)|
|[SheetBeforeDoubleClick](http://msdn.microsoft.com/library/workbook-sheetbeforedoubleclick-event-excel%28Office.15%29.aspx)|
|[SheetBeforeRightClick](http://msdn.microsoft.com/library/workbook-sheetbeforerightclick-event-excel%28Office.15%29.aspx)|
|[SheetCalculate](http://msdn.microsoft.com/library/workbook-sheetcalculate-event-excel%28Office.15%29.aspx)|
|[SheetChange](http://msdn.microsoft.com/library/workbook-sheetchange-event-excel%28Office.15%29.aspx)|
|[SheetDeactivate](http://msdn.microsoft.com/library/workbook-sheetdeactivate-event-excel%28Office.15%29.aspx)|
|[SheetFollowHyperlink](http://msdn.microsoft.com/library/workbook-sheetfollowhyperlink-event-excel%28Office.15%29.aspx)|
|[SheetLensGalleryRenderComplete](http://msdn.microsoft.com/library/workbook-sheetlensgalleryrendercomplete-event-excel%28Office.15%29.aspx)|
|[SheetPivotTableAfterValueChange](http://msdn.microsoft.com/library/workbook-sheetpivottableaftervaluechange-event-excel%28Office.15%29.aspx)|
|[SheetPivotTableBeforeAllocateChanges](http://msdn.microsoft.com/library/workbook-sheetpivottablebeforeallocatechanges-event-excel%28Office.15%29.aspx)|
|[SheetPivotTableBeforeCommitChanges](http://msdn.microsoft.com/library/workbook-sheetpivottablebeforecommitchanges-event-excel%28Office.15%29.aspx)|
|[SheetPivotTableBeforeDiscardChanges](http://msdn.microsoft.com/library/workbook-sheetpivottablebeforediscardchanges-event-excel%28Office.15%29.aspx)|
|[SheetPivotTableChangeSync](http://msdn.microsoft.com/library/workbook-sheetpivottablechangesync-event-excel%28Office.15%29.aspx)|
|[SheetPivotTableUpdate](http://msdn.microsoft.com/library/workbook-sheetpivottableupdate-event-excel%28Office.15%29.aspx)|
|[SheetSelectionChange](http://msdn.microsoft.com/library/workbook-sheetselectionchange-event-excel%28Office.15%29.aspx)|
|[SheetTableUpdate](http://msdn.microsoft.com/library/workbook-sheettableupdate-event-excel%28Office.15%29.aspx)|
|[Sync](http://msdn.microsoft.com/library/workbook-sync-event-excel%28Office.15%29.aspx)|
|[WindowActivate](http://msdn.microsoft.com/library/workbook-windowactivate-event-excel%28Office.15%29.aspx)|
|[WindowDeactivate](http://msdn.microsoft.com/library/workbook-windowdeactivate-event-excel%28Office.15%29.aspx)|
|[WindowResize](http://msdn.microsoft.com/library/workbook-windowresize-event-excel%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[AcceptAllChanges](http://msdn.microsoft.com/library/workbook-acceptallchanges-method-excel%28Office.15%29.aspx)|
|[Activate](http://msdn.microsoft.com/library/workbook-activate-method-excel%28Office.15%29.aspx)|
|[AddToFavorites](http://msdn.microsoft.com/library/workbook-addtofavorites-method-excel%28Office.15%29.aspx)|
|[ApplyTheme](http://msdn.microsoft.com/library/workbook-applytheme-method-excel%28Office.15%29.aspx)|
|[BreakLink](http://msdn.microsoft.com/library/workbook-breaklink-method-excel%28Office.15%29.aspx)|
|[CanCheckIn](http://msdn.microsoft.com/library/workbook-cancheckin-method-excel%28Office.15%29.aspx)|
|[ChangeFileAccess](http://msdn.microsoft.com/library/workbook-changefileaccess-method-excel%28Office.15%29.aspx)|
|[ChangeLink](http://msdn.microsoft.com/library/workbook-changelink-method-excel%28Office.15%29.aspx)|
|[CheckIn](http://msdn.microsoft.com/library/workbook-checkin-method-excel%28Office.15%29.aspx)|
|[CheckInWithVersion](http://msdn.microsoft.com/library/workbook-checkinwithversion-method-excel%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/workbook-close-method-excel%28Office.15%29.aspx)|
|[DeleteNumberFormat](http://msdn.microsoft.com/library/workbook-deletenumberformat-method-excel%28Office.15%29.aspx)|
|[EnableConnections](http://msdn.microsoft.com/library/workbook-enableconnections-method-excel%28Office.15%29.aspx)|
|[EndReview](http://msdn.microsoft.com/library/workbook-endreview-method-excel%28Office.15%29.aspx)|
|[ExclusiveAccess](http://msdn.microsoft.com/library/workbook-exclusiveaccess-method-excel%28Office.15%29.aspx)|
|[ExportAsFixedFormat](http://msdn.microsoft.com/library/workbook-exportasfixedformat-method-excel%28Office.15%29.aspx)|
|[FollowHyperlink](http://msdn.microsoft.com/library/workbook-followhyperlink-method-excel%28Office.15%29.aspx)|
|[ForwardMailer](http://msdn.microsoft.com/library/workbook-forwardmailer-method-excel%28Office.15%29.aspx)|
|[GetWorkflowTasks](http://msdn.microsoft.com/library/workbook-getworkflowtasks-method-excel%28Office.15%29.aspx)|
|[GetWorkflowTemplates](http://msdn.microsoft.com/library/workbook-getworkflowtemplates-method-excel%28Office.15%29.aspx)|
|[HighlightChangesOptions](http://msdn.microsoft.com/library/workbook-highlightchangesoptions-method-excel%28Office.15%29.aspx)|
|[LinkInfo](http://msdn.microsoft.com/library/workbook-linkinfo-method-excel%28Office.15%29.aspx)|
|[LinkSources](http://msdn.microsoft.com/library/workbook-linksources-method-excel%28Office.15%29.aspx)|
|[LockServerFile](http://msdn.microsoft.com/library/workbook-lockserverfile-method-excel%28Office.15%29.aspx)|
|[MergeWorkbook](http://msdn.microsoft.com/library/workbook-mergeworkbook-method-excel%28Office.15%29.aspx)|
|[NewWindow](http://msdn.microsoft.com/library/workbook-newwindow-method-excel%28Office.15%29.aspx)|
|[OpenLinks](http://msdn.microsoft.com/library/workbook-openlinks-method-excel%28Office.15%29.aspx)|
|[PivotCaches](http://msdn.microsoft.com/library/workbook-pivotcaches-method-excel%28Office.15%29.aspx)|
|[Post](http://msdn.microsoft.com/library/workbook-post-method-excel%28Office.15%29.aspx)|
|[PrintOut](http://msdn.microsoft.com/library/workbook-printout-method-excel%28Office.15%29.aspx)|
|[PrintPreview](http://msdn.microsoft.com/library/workbook-printpreview-method-excel%28Office.15%29.aspx)|
|[Protect](http://msdn.microsoft.com/library/workbook-protect-method-excel%28Office.15%29.aspx)|
|[ProtectSharing](http://msdn.microsoft.com/library/workbook-protectsharing-method-excel%28Office.15%29.aspx)|
|[PurgeChangeHistoryNow](http://msdn.microsoft.com/library/workbook-purgechangehistorynow-method-excel%28Office.15%29.aspx)|
|[RefreshAll](http://msdn.microsoft.com/library/workbook-refreshall-method-excel%28Office.15%29.aspx)|
|[RejectAllChanges](http://msdn.microsoft.com/library/workbook-rejectallchanges-method-excel%28Office.15%29.aspx)|
|[ReloadAs](http://msdn.microsoft.com/library/workbook-reloadas-method-excel%28Office.15%29.aspx)|
|[RemoveDocumentInformation](http://msdn.microsoft.com/library/workbook-removedocumentinformation-method-excel%28Office.15%29.aspx)|
|[RemoveUser](http://msdn.microsoft.com/library/workbook-removeuser-method-excel%28Office.15%29.aspx)|
|[Reply](http://msdn.microsoft.com/library/workbook-reply-method-excel%28Office.15%29.aspx)|
|[ReplyAll](http://msdn.microsoft.com/library/workbook-replyall-method-excel%28Office.15%29.aspx)|
|[ReplyWithChanges](http://msdn.microsoft.com/library/workbook-replywithchanges-method-excel%28Office.15%29.aspx)|
|[ResetColors](http://msdn.microsoft.com/library/workbook-resetcolors-method-excel%28Office.15%29.aspx)|
|[RunAutoMacros](http://msdn.microsoft.com/library/workbook-runautomacros-method-excel%28Office.15%29.aspx)|
|[Save](http://msdn.microsoft.com/library/workbook-save-method-excel%28Office.15%29.aspx)|
|[SaveAs](http://msdn.microsoft.com/library/workbook-saveas-method-excel%28Office.15%29.aspx)|
|[SaveAsXMLData](http://msdn.microsoft.com/library/workbook-saveasxmldata-method-excel%28Office.15%29.aspx)|
|[SaveCopyAs](http://msdn.microsoft.com/library/workbook-savecopyas-method-excel%28Office.15%29.aspx)|
|[SendFaxOverInternet](http://msdn.microsoft.com/library/workbook-sendfaxoverinternet-method-excel%28Office.15%29.aspx)|
|[SendForReview](http://msdn.microsoft.com/library/workbook-sendforreview-method-excel%28Office.15%29.aspx)|
|[SendMail](http://msdn.microsoft.com/library/workbook-sendmail-method-excel%28Office.15%29.aspx)|
|[SendMailer](http://msdn.microsoft.com/library/workbook-sendmailer-method-excel%28Office.15%29.aspx)|
|[SetLinkOnData](http://msdn.microsoft.com/library/workbook-setlinkondata-method-excel%28Office.15%29.aspx)|
|[SetPasswordEncryptionOptions](http://msdn.microsoft.com/library/workbook-setpasswordencryptionoptions-method-excel%28Office.15%29.aspx)|
|[ToggleFormsDesign](http://msdn.microsoft.com/library/workbook-toggleformsdesign-method-excel%28Office.15%29.aspx)|
|[Unprotect](http://msdn.microsoft.com/library/workbook-unprotect-method-excel%28Office.15%29.aspx)|
|[UnprotectSharing](http://msdn.microsoft.com/library/workbook-unprotectsharing-method-excel%28Office.15%29.aspx)|
|[UpdateFromFile](http://msdn.microsoft.com/library/workbook-updatefromfile-method-excel%28Office.15%29.aspx)|
|[UpdateLink](http://msdn.microsoft.com/library/workbook-updatelink-method-excel%28Office.15%29.aspx)|
|[WebPagePreview](http://msdn.microsoft.com/library/workbook-webpagepreview-method-excel%28Office.15%29.aspx)|
|[XmlImport](http://msdn.microsoft.com/library/workbook-xmlimport-method-excel%28Office.15%29.aspx)|
|[XmlImportXml](http://msdn.microsoft.com/library/workbook-xmlimportxml-method-excel%28Office.15%29.aspx)|
|[CreateForecastSheet](http://msdn.microsoft.com/library/workbook-createforecastsheet-method-excel%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AccuracyVersion](http://msdn.microsoft.com/library/workbook-accuracyversion-property-excel%28Office.15%29.aspx)|
|[ActiveChart](http://msdn.microsoft.com/library/workbook-activechart-property-excel%28Office.15%29.aspx)|
|[ActiveSheet](http://msdn.microsoft.com/library/workbook-activesheet-property-excel%28Office.15%29.aspx)|
|[ActiveSlicer](http://msdn.microsoft.com/library/workbook-activeslicer-property-excel%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/workbook-application-property-excel%28Office.15%29.aspx)|
|[AutoUpdateFrequency](http://msdn.microsoft.com/library/workbook-autoupdatefrequency-property-excel%28Office.15%29.aspx)|
|[AutoUpdateSaveChanges](http://msdn.microsoft.com/library/workbook-autoupdatesavechanges-property-excel%28Office.15%29.aspx)|
|[BuiltinDocumentProperties](http://msdn.microsoft.com/library/workbook-builtindocumentproperties-property-excel%28Office.15%29.aspx)|
|[CalculationVersion](http://msdn.microsoft.com/library/workbook-calculationversion-property-excel%28Office.15%29.aspx)|
|[CaseSensitive](http://msdn.microsoft.com/library/workbook-casesensitive-property-excel%28Office.15%29.aspx)|
|[ChangeHistoryDuration](http://msdn.microsoft.com/library/workbook-changehistoryduration-property-excel%28Office.15%29.aspx)|
|[ChartDataPointTrack](http://msdn.microsoft.com/library/workbook-chartdatapointtrack-property-excel%28Office.15%29.aspx)|
|[Charts](http://msdn.microsoft.com/library/workbook-charts-property-excel%28Office.15%29.aspx)|
|[CheckCompatibility](http://msdn.microsoft.com/library/workbook-checkcompatibility-property-excel%28Office.15%29.aspx)|
|[CodeName](http://msdn.microsoft.com/library/workbook-codename-property-excel%28Office.15%29.aspx)|
|[Colors](http://msdn.microsoft.com/library/workbook-colors-property-excel%28Office.15%29.aspx)|
|[CommandBars](http://msdn.microsoft.com/library/workbook-commandbars-property-excel%28Office.15%29.aspx)|
|[ConflictResolution](http://msdn.microsoft.com/library/workbook-conflictresolution-property-excel%28Office.15%29.aspx)|
|[Connections](http://msdn.microsoft.com/library/workbook-connections-property-excel%28Office.15%29.aspx)|
|[ConnectionsDisabled](http://msdn.microsoft.com/library/workbook-connectionsdisabled-property-excel%28Office.15%29.aspx)|
|[Container](http://msdn.microsoft.com/library/workbook-container-property-excel%28Office.15%29.aspx)|
|[ContentTypeProperties](http://msdn.microsoft.com/library/workbook-contenttypeproperties-property-excel%28Office.15%29.aspx)|
|[CreateBackup](http://msdn.microsoft.com/library/workbook-createbackup-property-excel%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/workbook-creator-property-excel%28Office.15%29.aspx)|
|[CustomDocumentProperties](http://msdn.microsoft.com/library/workbook-customdocumentproperties-property-excel%28Office.15%29.aspx)|
|[CustomViews](http://msdn.microsoft.com/library/workbook-customviews-property-excel%28Office.15%29.aspx)|
|[CustomXMLParts](http://msdn.microsoft.com/library/workbook-customxmlparts-property-excel%28Office.15%29.aspx)|
|[Date1904](http://msdn.microsoft.com/library/workbook-date1904-property-excel%28Office.15%29.aspx)|
|[DefaultPivotTableStyle](http://msdn.microsoft.com/library/workbook-defaultpivottablestyle-property-excel%28Office.15%29.aspx)|
|[DefaultSlicerStyle](http://msdn.microsoft.com/library/workbook-defaultslicerstyle-property-excel%28Office.15%29.aspx)|
|[DefaultTableStyle](http://msdn.microsoft.com/library/workbook-defaulttablestyle-property-excel%28Office.15%29.aspx)|
|[DefaultTimelineStyle](http://msdn.microsoft.com/library/workbook-defaulttimelinestyle-property-excel%28Office.15%29.aspx)|
|[DisplayDrawingObjects](http://msdn.microsoft.com/library/workbook-displaydrawingobjects-property-excel%28Office.15%29.aspx)|
|[DisplayInkComments](http://msdn.microsoft.com/library/workbook-displayinkcomments-property-excel%28Office.15%29.aspx)|
|[DocumentInspectors](http://msdn.microsoft.com/library/workbook-documentinspectors-property-excel%28Office.15%29.aspx)|
|[DocumentLibraryVersions](http://msdn.microsoft.com/library/workbook-documentlibraryversions-property-excel%28Office.15%29.aspx)|
|[DoNotPromptForConvert](http://msdn.microsoft.com/library/workbook-donotpromptforconvert-property-excel%28Office.15%29.aspx)|
|[EnableAutoRecover](http://msdn.microsoft.com/library/workbook-enableautorecover-property-excel%28Office.15%29.aspx)|
|[EncryptionProvider](http://msdn.microsoft.com/library/workbook-encryptionprovider-property-excel%28Office.15%29.aspx)|
|[EnvelopeVisible](http://msdn.microsoft.com/library/workbook-envelopevisible-property-excel%28Office.15%29.aspx)|
|[Excel4IntlMacroSheets](http://msdn.microsoft.com/library/workbook-excel4intlmacrosheets-property-excel%28Office.15%29.aspx)|
|[Excel4MacroSheets](http://msdn.microsoft.com/library/workbook-excel4macrosheets-property-excel%28Office.15%29.aspx)|
|[Excel8CompatibilityMode](http://msdn.microsoft.com/library/workbook-excel8compatibilitymode-property-excel%28Office.15%29.aspx)|
|[FileFormat](http://msdn.microsoft.com/library/workbook-fileformat-property-excel%28Office.15%29.aspx)|
|[Final](http://msdn.microsoft.com/library/workbook-final-property-excel%28Office.15%29.aspx)|
|[ForceFullCalculation](http://msdn.microsoft.com/library/workbook-forcefullcalculation-property-excel%28Office.15%29.aspx)|
|[FullName](http://msdn.microsoft.com/library/workbook-fullname-property-excel%28Office.15%29.aspx)|
|[FullNameURLEncoded](http://msdn.microsoft.com/library/workbook-fullnameurlencoded-property-excel%28Office.15%29.aspx)|
|[HasPassword](http://msdn.microsoft.com/library/workbook-haspassword-property-excel%28Office.15%29.aspx)|
|[HasVBProject](http://msdn.microsoft.com/library/workbook-hasvbproject-property-excel%28Office.15%29.aspx)|
|[HighlightChangesOnScreen](http://msdn.microsoft.com/library/workbook-highlightchangesonscreen-property-excel%28Office.15%29.aspx)|
|[IconSets](http://msdn.microsoft.com/library/workbook-iconsets-property-excel%28Office.15%29.aspx)|
|[InactiveListBorderVisible](http://msdn.microsoft.com/library/workbook-inactivelistbordervisible-property-excel%28Office.15%29.aspx)|
|[IsAddin](http://msdn.microsoft.com/library/workbook-isaddin-property-excel%28Office.15%29.aspx)|
|[IsInplace](http://msdn.microsoft.com/library/workbook-isinplace-property-excel%28Office.15%29.aspx)|
|[KeepChangeHistory](http://msdn.microsoft.com/library/workbook-keepchangehistory-property-excel%28Office.15%29.aspx)|
|[ListChangesOnNewSheet](http://msdn.microsoft.com/library/workbook-listchangesonnewsheet-property-excel%28Office.15%29.aspx)|
|[Mailer](http://msdn.microsoft.com/library/workbook-mailer-property-excel%28Office.15%29.aspx)|
|[Model](http://msdn.microsoft.com/library/workbook-model-property-excel%28Office.15%29.aspx)|
|[MultiUserEditing](http://msdn.microsoft.com/library/workbook-multiuserediting-property-excel%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/workbook-name-property-excel%28Office.15%29.aspx)|
|[Names](http://msdn.microsoft.com/library/workbook-names-property-excel%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/workbook-parent-property-excel%28Office.15%29.aspx)|
|[Password](http://msdn.microsoft.com/library/workbook-password-property-excel%28Office.15%29.aspx)|
|[PasswordEncryptionAlgorithm](http://msdn.microsoft.com/library/workbook-passwordencryptionalgorithm-property-excel%28Office.15%29.aspx)|
|[PasswordEncryptionFileProperties](http://msdn.microsoft.com/library/workbook-passwordencryptionfileproperties-property-excel%28Office.15%29.aspx)|
|[PasswordEncryptionKeyLength](http://msdn.microsoft.com/library/workbook-passwordencryptionkeylength-property-excel%28Office.15%29.aspx)|
|[PasswordEncryptionProvider](http://msdn.microsoft.com/library/workbook-passwordencryptionprovider-property-excel%28Office.15%29.aspx)|
|[Path](http://msdn.microsoft.com/library/workbook-path-property-excel%28Office.15%29.aspx)|
|[Permission](http://msdn.microsoft.com/library/workbook-permission-property-excel%28Office.15%29.aspx)|
|[PersonalViewListSettings](http://msdn.microsoft.com/library/workbook-personalviewlistsettings-property-excel%28Office.15%29.aspx)|
|[PersonalViewPrintSettings](http://msdn.microsoft.com/library/workbook-personalviewprintsettings-property-excel%28Office.15%29.aspx)|
|[PivotTables](http://msdn.microsoft.com/library/workbook-pivottables-property-excel%28Office.15%29.aspx)|
|[PrecisionAsDisplayed](http://msdn.microsoft.com/library/workbook-precisionasdisplayed-property-excel%28Office.15%29.aspx)|
|[ProtectStructure](http://msdn.microsoft.com/library/workbook-protectstructure-property-excel%28Office.15%29.aspx)|
|[ProtectWindows](http://msdn.microsoft.com/library/workbook-protectwindows-property-excel%28Office.15%29.aspx)|
|[PublishObjects](http://msdn.microsoft.com/library/workbook-publishobjects-property-excel%28Office.15%29.aspx)|
|[ReadOnly](http://msdn.microsoft.com/library/workbook-readonly-property-excel%28Office.15%29.aspx)|
|[ReadOnlyRecommended](http://msdn.microsoft.com/library/workbook-readonlyrecommended-property-excel%28Office.15%29.aspx)|
|[RemovePersonalInformation](http://msdn.microsoft.com/library/workbook-removepersonalinformation-property-excel%28Office.15%29.aspx)|
|[Research](http://msdn.microsoft.com/library/workbook-research-property-excel%28Office.15%29.aspx)|
|[RevisionNumber](http://msdn.microsoft.com/library/workbook-revisionnumber-property-excel%28Office.15%29.aspx)|
|[Saved](http://msdn.microsoft.com/library/workbook-saved-property-excel%28Office.15%29.aspx)|
|[SaveLinkValues](http://msdn.microsoft.com/library/workbook-savelinkvalues-property-excel%28Office.15%29.aspx)|
|[ServerPolicy](http://msdn.microsoft.com/library/workbook-serverpolicy-property-excel%28Office.15%29.aspx)|
|[ServerViewableItems](http://msdn.microsoft.com/library/workbook-serverviewableitems-property-excel%28Office.15%29.aspx)|
|[SharedWorkspace](http://msdn.microsoft.com/library/workbook-sharedworkspace-property-excel%28Office.15%29.aspx)|
|[Sheets](http://msdn.microsoft.com/library/workbook-sheets-property-excel%28Office.15%29.aspx)|
|[ShowConflictHistory](http://msdn.microsoft.com/library/workbook-showconflicthistory-property-excel%28Office.15%29.aspx)|
|[ShowPivotChartActiveFields](http://msdn.microsoft.com/library/workbook-showpivotchartactivefields-property-excel%28Office.15%29.aspx)|
|[ShowPivotTableFieldList](http://msdn.microsoft.com/library/workbook-showpivottablefieldlist-property-excel%28Office.15%29.aspx)|
|[Signatures](http://msdn.microsoft.com/library/workbook-signatures-property-excel%28Office.15%29.aspx)|
|[SlicerCaches](http://msdn.microsoft.com/library/workbook-slicercaches-property-excel%28Office.15%29.aspx)|
|[SmartDocument](http://msdn.microsoft.com/library/workbook-smartdocument-property-excel%28Office.15%29.aspx)|
|[Styles](http://msdn.microsoft.com/library/workbook-styles-property-excel%28Office.15%29.aspx)|
|[Sync](http://msdn.microsoft.com/library/workbook-sync-property-excel%28Office.15%29.aspx)|
|[TableStyles](http://msdn.microsoft.com/library/workbook-tablestyles-property-excel%28Office.15%29.aspx)|
|[TemplateRemoveExtData](http://msdn.microsoft.com/library/workbook-templateremoveextdata-property-excel%28Office.15%29.aspx)|
|[Theme](http://msdn.microsoft.com/library/workbook-theme-property-excel%28Office.15%29.aspx)|
|[UpdateLinks](http://msdn.microsoft.com/library/workbook-updatelinks-property-excel%28Office.15%29.aspx)|
|[UpdateRemoteReferences](http://msdn.microsoft.com/library/workbook-updateremotereferences-property-excel%28Office.15%29.aspx)|
|[UserStatus](http://msdn.microsoft.com/library/workbook-userstatus-property-excel%28Office.15%29.aspx)|
|[UseWholeCellCriteria](http://msdn.microsoft.com/library/workbook-usewholecellcriteria-property-excel%28Office.15%29.aspx)|
|[UseWildcards](http://msdn.microsoft.com/library/workbook-usewildcards-property-excel%28Office.15%29.aspx)|
|[VBASigned](http://msdn.microsoft.com/library/workbook-vbasigned-property-excel%28Office.15%29.aspx)|
|[VBProject](http://msdn.microsoft.com/library/workbook-vbproject-property-excel%28Office.15%29.aspx)|
|[WebOptions](http://msdn.microsoft.com/library/workbook-weboptions-property-excel%28Office.15%29.aspx)|
|[Windows](http://msdn.microsoft.com/library/workbook-windows-property-excel%28Office.15%29.aspx)|
|[Worksheets](http://msdn.microsoft.com/library/workbook-worksheets-property-excel%28Office.15%29.aspx)|
|[WritePassword](http://msdn.microsoft.com/library/workbook-writepassword-property-excel%28Office.15%29.aspx)|
|[WriteReserved](http://msdn.microsoft.com/library/workbook-writereserved-property-excel%28Office.15%29.aspx)|
|[WriteReservedBy](http://msdn.microsoft.com/library/workbook-writereservedby-property-excel%28Office.15%29.aspx)|
|[XmlMaps](http://msdn.microsoft.com/library/workbook-xmlmaps-property-excel%28Office.15%29.aspx)|
|[XmlNamespaces](http://msdn.microsoft.com/library/workbook-xmlnamespaces-property-excel%28Office.15%29.aspx)|
|[Queries](http://msdn.microsoft.com/library/workbook-queries-property-excel%28Office.15%29.aspx)|

## About the Contributor
<a name="AboutContributor"> </a>

Holy Macro! Books publishes entertaining books for people who use Microsoft Office. See the complete catalog at MrExcel.com. 


## See also
<a name="AboutContributor"> </a>


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/object-model-excel-vba-reference%28Office.15%29.aspx)
<<<<<<< HEAD
=======

>>>>>>> d7667e83d23dbf8ebf5bf068ba6fed14c840c0f5

