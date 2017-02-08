---
title: Application Object (Excel)
keywords: vbaxl10.chm182073
f1_keywords:
- vbaxl10.chm182073
ms.prod: EXCEL
api_name:
- Excel.Application
ms.assetid: 19b73597-5cf9-4f56-8227-b5211f657f6f
---


# Application Object (Excel)

Represents the entire Microsoft Excel application.


## Example

Use the  **Application** property to return the **Application** object. The following example applies the **Windows** property to the **Application** object.


```
Application.Windows("book1.xls").Activate
```

The following example creates a Microsoft Excel workbook object in another application and then opens a workbook in Microsoft Excel.




```
Set xl = CreateObject("Excel.Sheet") 
xl.Application.Workbooks.Open "newbook.xls"
```

Many of the properties and methods that return the most common user-interface objects, such as the active cell ( **ActiveCell** property), can be used without the **Application** object qualifier. For example, instead of writing




```
Application.ActiveCell.Font.Bold = True
```

You can write 




```
ActiveCell.Font.Bold = True
```


## Remarks

The  **Application** object contains:


- Application-wide settings and options.
    
- Methods that return top-level objects, such as  **[ActiveCell](http://msdn.microsoft.com/library/application-activecell-property-excel%28Office.15%29.aspx)**, **[ActiveSheet](http://msdn.microsoft.com/library/application-activesheet-property-excel%28Office.15%29.aspx)**, and so on.
    



## Events



|**Name**|
|:-----|
|[AfterCalculate](http://msdn.microsoft.com/library/application-aftercalculate-event-excel%28Office.15%29.aspx)|
|[NewWorkbook](http://msdn.microsoft.com/library/application-newworkbook-event-excel%28Office.15%29.aspx)|
|[ProtectedViewWindowActivate](http://msdn.microsoft.com/library/application-protectedviewwindowactivate-event-excel%28Office.15%29.aspx)|
|[ProtectedViewWindowBeforeClose](http://msdn.microsoft.com/library/application-protectedviewwindowbeforeclose-event-excel%28Office.15%29.aspx)|
|[ProtectedViewWindowBeforeEdit](http://msdn.microsoft.com/library/application-protectedviewwindowbeforeedit-event-excel%28Office.15%29.aspx)|
|[ProtectedViewWindowDeactivate](http://msdn.microsoft.com/library/application-protectedviewwindowdeactivate-event-excel%28Office.15%29.aspx)|
|[ProtectedViewWindowOpen](http://msdn.microsoft.com/library/application-protectedviewwindowopen-event-excel%28Office.15%29.aspx)|
|[ProtectedViewWindowResize](http://msdn.microsoft.com/library/application-protectedviewwindowresize-event-excel%28Office.15%29.aspx)|
|[SheetActivate](http://msdn.microsoft.com/library/application-sheetactivate-event-excel%28Office.15%29.aspx)|
|[SheetBeforeDelete](http://msdn.microsoft.com/library/application-sheetbeforedelete-event-excel%28Office.15%29.aspx)|
|[SheetBeforeDoubleClick](http://msdn.microsoft.com/library/application-sheetbeforedoubleclick-event-excel%28Office.15%29.aspx)|
|[SheetBeforeRightClick](http://msdn.microsoft.com/library/application-sheetbeforerightclick-event-excel%28Office.15%29.aspx)|
|[SheetCalculate](http://msdn.microsoft.com/library/application-sheetcalculate-event-excel%28Office.15%29.aspx)|
|[SheetChange](http://msdn.microsoft.com/library/application-sheetchange-event-excel%28Office.15%29.aspx)|
|[SheetDeactivate](http://msdn.microsoft.com/library/application-sheetdeactivate-event-excel%28Office.15%29.aspx)|
|[SheetFollowHyperlink](http://msdn.microsoft.com/library/application-sheetfollowhyperlink-event-excel%28Office.15%29.aspx)|
|[SheetLensGalleryRenderComplete](http://msdn.microsoft.com/library/application-sheetlensgalleryrendercomplete-event-excel%28Office.15%29.aspx)|
|[SheetPivotTableAfterValueChange](http://msdn.microsoft.com/library/application-sheetpivottableaftervaluechange-event-excel%28Office.15%29.aspx)|
|[SheetPivotTableBeforeAllocateChanges](http://msdn.microsoft.com/library/application-sheetpivottablebeforeallocatechanges-event-excel%28Office.15%29.aspx)|
|[SheetPivotTableBeforeCommitChanges](http://msdn.microsoft.com/library/application-sheetpivottablebeforecommitchanges-event-excel%28Office.15%29.aspx)|
|[SheetPivotTableBeforeDiscardChanges](http://msdn.microsoft.com/library/application-sheetpivottablebeforediscardchanges-event-excel%28Office.15%29.aspx)|
|[SheetPivotTableUpdate](http://msdn.microsoft.com/library/application-sheetpivottableupdate-event-excel%28Office.15%29.aspx)|
|[SheetSelectionChange](http://msdn.microsoft.com/library/application-sheetselectionchange-event-excel%28Office.15%29.aspx)|
|[SheetTableUpdate](http://msdn.microsoft.com/library/application-sheettableupdate-event-excel%28Office.15%29.aspx)|
|[WindowActivate](http://msdn.microsoft.com/library/application-windowactivate-event-excel%28Office.15%29.aspx)|
|[WindowDeactivate](http://msdn.microsoft.com/library/application-windowdeactivate-event-excel%28Office.15%29.aspx)|
|[WindowResize](http://msdn.microsoft.com/library/application-windowresize-event-excel%28Office.15%29.aspx)|
|[WorkbookActivate](http://msdn.microsoft.com/library/application-workbookactivate-event-excel%28Office.15%29.aspx)|
|[WorkbookAddinInstall](http://msdn.microsoft.com/library/application-workbookaddininstall-event-excel%28Office.15%29.aspx)|
|[WorkbookAddinUninstall](http://msdn.microsoft.com/library/application-workbookaddinuninstall-event-excel%28Office.15%29.aspx)|
|[WorkbookAfterSave](http://msdn.microsoft.com/library/application-workbookaftersave-event-excel%28Office.15%29.aspx)|
|[WorkbookAfterXmlExport](http://msdn.microsoft.com/library/application-workbookafterxmlexport-event-excel%28Office.15%29.aspx)|
|[WorkbookAfterXmlImport](http://msdn.microsoft.com/library/application-workbookafterxmlimport-event-excel%28Office.15%29.aspx)|
|[WorkbookBeforeClose](http://msdn.microsoft.com/library/application-workbookbeforeclose-event-excel%28Office.15%29.aspx)|
|[WorkbookBeforePrint](http://msdn.microsoft.com/library/application-workbookbeforeprint-event-excel%28Office.15%29.aspx)|
|[WorkbookBeforeSave](http://msdn.microsoft.com/library/application-workbookbeforesave-event-excel%28Office.15%29.aspx)|
|[WorkbookBeforeXmlExport](http://msdn.microsoft.com/library/application-workbookbeforexmlexport-event-excel%28Office.15%29.aspx)|
|[WorkbookBeforeXmlImport](http://msdn.microsoft.com/library/application-workbookbeforexmlimport-event-excel%28Office.15%29.aspx)|
|[WorkbookDeactivate](http://msdn.microsoft.com/library/application-workbookdeactivate-event-excel%28Office.15%29.aspx)|
|[WorkbookModelChange](http://msdn.microsoft.com/library/application-workbookmodelchange-event-excel%28Office.15%29.aspx)|
|[WorkbookNewChart](http://msdn.microsoft.com/library/application-workbooknewchart-event-excel%28Office.15%29.aspx)|
|[WorkbookNewSheet](http://msdn.microsoft.com/library/application-workbooknewsheet-event-excel%28Office.15%29.aspx)|
|[WorkbookOpen](http://msdn.microsoft.com/library/application-workbookopen-event-excel%28Office.15%29.aspx)|
|[WorkbookPivotTableCloseConnection](http://msdn.microsoft.com/library/application-workbookpivottablecloseconnection-event-excel%28Office.15%29.aspx)|
|[WorkbookPivotTableOpenConnection](http://msdn.microsoft.com/library/application-workbookpivottableopenconnection-event-excel%28Office.15%29.aspx)|
|[WorkbookRowsetComplete](http://msdn.microsoft.com/library/application-workbookrowsetcomplete-event-excel%28Office.15%29.aspx)|
|[WorkbookSync](http://msdn.microsoft.com/library/application-workbooksync-event-excel%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[ActivateMicrosoftApp](http://msdn.microsoft.com/library/application-activatemicrosoftapp-method-excel%28Office.15%29.aspx)|
|[AddCustomList](http://msdn.microsoft.com/library/application-addcustomlist-method-excel%28Office.15%29.aspx)|
|[Calculate](http://msdn.microsoft.com/library/application-calculate-method-excel%28Office.15%29.aspx)|
|[CalculateFull](http://msdn.microsoft.com/library/application-calculatefull-method-excel%28Office.15%29.aspx)|
|[CalculateFullRebuild](http://msdn.microsoft.com/library/application-calculatefullrebuild-method-excel%28Office.15%29.aspx)|
|[CalculateUntilAsyncQueriesDone](http://msdn.microsoft.com/library/application-calculateuntilasyncqueriesdone-method-excel%28Office.15%29.aspx)|
|[CentimetersToPoints](http://msdn.microsoft.com/library/application-centimeterstopoints-method-excel%28Office.15%29.aspx)|
|[CheckAbort](http://msdn.microsoft.com/library/application-checkabort-method-excel%28Office.15%29.aspx)|
|[CheckSpelling](http://msdn.microsoft.com/library/application-checkspelling-method-excel%28Office.15%29.aspx)|
|[ConvertFormula](http://msdn.microsoft.com/library/application-convertformula-method-excel%28Office.15%29.aspx)|
|[DDEExecute](http://msdn.microsoft.com/library/application-ddeexecute-method-excel%28Office.15%29.aspx)|
|[DDEInitiate](http://msdn.microsoft.com/library/application-ddeinitiate-method-excel%28Office.15%29.aspx)|
|[DDEPoke](http://msdn.microsoft.com/library/application-ddepoke-method-excel%28Office.15%29.aspx)|
|[DDERequest](http://msdn.microsoft.com/library/application-dderequest-method-excel%28Office.15%29.aspx)|
|[DDETerminate](http://msdn.microsoft.com/library/application-ddeterminate-method-excel%28Office.15%29.aspx)|
|[DeleteCustomList](http://msdn.microsoft.com/library/application-deletecustomlist-method-excel%28Office.15%29.aspx)|
|[DisplayXMLSourcePane](http://msdn.microsoft.com/library/application-displayxmlsourcepane-method-excel%28Office.15%29.aspx)|
|[DoubleClick](http://msdn.microsoft.com/library/application-doubleclick-method-excel%28Office.15%29.aspx)|
|[Evaluate](http://msdn.microsoft.com/library/application-evaluate-method-excel%28Office.15%29.aspx)|
|[ExecuteExcel4Macro](http://msdn.microsoft.com/library/application-executeexcel4macro-method-excel%28Office.15%29.aspx)|
|[FindFile](http://msdn.microsoft.com/library/application-findfile-method-excel%28Office.15%29.aspx)|
|[GetCustomListContents](http://msdn.microsoft.com/library/application-getcustomlistcontents-method-excel%28Office.15%29.aspx)|
|[GetCustomListNum](http://msdn.microsoft.com/library/application-getcustomlistnum-method-excel%28Office.15%29.aspx)|
|[GetOpenFilename](http://msdn.microsoft.com/library/application-getopenfilename-method-excel%28Office.15%29.aspx)|
|[GetPhonetic](http://msdn.microsoft.com/library/application-getphonetic-method-excel%28Office.15%29.aspx)|
|[GetSaveAsFilename](http://msdn.microsoft.com/library/application-getsaveasfilename-method-excel%28Office.15%29.aspx)|
|[Goto](http://msdn.microsoft.com/library/application-goto-method-excel%28Office.15%29.aspx)|
|[Help](http://msdn.microsoft.com/library/application-help-method-excel%28Office.15%29.aspx)|
|[InchesToPoints](http://msdn.microsoft.com/library/application-inchestopoints-method-excel%28Office.15%29.aspx)|
|[InputBox](http://msdn.microsoft.com/library/application-inputbox-method-excel%28Office.15%29.aspx)|
|[Intersect](http://msdn.microsoft.com/library/application-intersect-method-excel%28Office.15%29.aspx)|
|[MacroOptions](http://msdn.microsoft.com/library/application-macrooptions-method-excel%28Office.15%29.aspx)|
|[MailLogoff](http://msdn.microsoft.com/library/application-maillogoff-method-excel%28Office.15%29.aspx)|
|[MailLogon](http://msdn.microsoft.com/library/application-maillogon-method-excel%28Office.15%29.aspx)|
|[NextLetter](http://msdn.microsoft.com/library/application-nextletter-method-excel%28Office.15%29.aspx)|
|[OnKey](http://msdn.microsoft.com/library/application-onkey-method-excel%28Office.15%29.aspx)||
|[OnRepeat](http://msdn.microsoft.com/library/application-onrepeat-method-excel%28Office.15%29.aspx)||
|[OnTime](http://msdn.microsoft.com/library/application-ontime-method-excel%28Office.15%29.aspx)||
|[OnUndo](http://msdn.microsoft.com/library/application-onundo-method-excel%28Office.15%29.aspx)||
|[Quit](http://msdn.microsoft.com/library/application-quit-method-excel%28Office.15%29.aspx)||
|[RecordMacro](http://msdn.microsoft.com/library/application-recordmacro-method-excel%28Office.15%29.aspx)||
|[RegisterXLL](http://msdn.microsoft.com/library/application-registerxll-method-excel%28Office.15%29.aspx)||
|[Repeat](http://msdn.microsoft.com/library/application-repeat-method-excel%28Office.15%29.aspx)||
|[Run](http://msdn.microsoft.com/library/application-run-method-excel%28Office.15%29.aspx)||
|[SendKeys](http://msdn.microsoft.com/library/application-sendkeys-method-excel%28Office.15%29.aspx)||
|[SharePointVersion](http://msdn.microsoft.com/library/application-sharepointversion-method-excel%28Office.15%29.aspx)||
|[Undo](http://msdn.microsoft.com/library/application-undo-method-excel%28Office.15%29.aspx)||
|[Union](http://msdn.microsoft.com/library/application-union-method-excel%28Office.15%29.aspx)||
|[Volatile](http://msdn.microsoft.com/library/application-volatile-method-excel%28Office.15%29.aspx)||
|[Wait](http://msdn.microsoft.com/library/application-wait-method-excel%28Office.15%29.aspx)||

## Properties



|**Name**|
|:-----|
|[ActiveCell](http://msdn.microsoft.com/library/application-activecell-property-excel%28Office.15%29.aspx)|
|[ActiveChart](http://msdn.microsoft.com/library/application-activechart-property-excel%28Office.15%29.aspx)|
|[ActiveEncryptionSession](http://msdn.microsoft.com/library/application-activeencryptionsession-property-excel%28Office.15%29.aspx)|
|[ActivePrinter](http://msdn.microsoft.com/library/application-activeprinter-property-excel%28Office.15%29.aspx)|
|[ActiveProtectedViewWindow](http://msdn.microsoft.com/library/application-activeprotectedviewwindow-property-excel%28Office.15%29.aspx)|
|[ActiveSheet](http://msdn.microsoft.com/library/application-activesheet-property-excel%28Office.15%29.aspx)|
|[ActiveWindow](http://msdn.microsoft.com/library/application-activewindow-property-excel%28Office.15%29.aspx)|
|[ActiveWorkbook](http://msdn.microsoft.com/library/application-activeworkbook-property-excel%28Office.15%29.aspx)|
|[AddIns](http://msdn.microsoft.com/library/application-addins-property-excel%28Office.15%29.aspx)|
|[AddIns2](http://msdn.microsoft.com/library/application-addins2-property-excel%28Office.15%29.aspx)|
|[AlertBeforeOverwriting](http://msdn.microsoft.com/library/application-alertbeforeoverwriting-property-excel%28Office.15%29.aspx)|
|[AltStartupPath](http://msdn.microsoft.com/library/application-altstartuppath-property-excel%28Office.15%29.aspx)|
|[AlwaysUseClearType](http://msdn.microsoft.com/library/application-alwaysusecleartype-property-excel%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/application-application-property-excel%28Office.15%29.aspx)|
|[ArbitraryXMLSupportAvailable](http://msdn.microsoft.com/library/application-arbitraryxmlsupportavailable-property-excel%28Office.15%29.aspx)|
|[AskToUpdateLinks](http://msdn.microsoft.com/library/application-asktoupdatelinks-property-excel%28Office.15%29.aspx)|
|[Assistance](http://msdn.microsoft.com/library/application-assistance-property-excel%28Office.15%29.aspx)|
|[AutoCorrect](http://msdn.microsoft.com/library/application-autocorrect-property-excel%28Office.15%29.aspx)|
|[AutoFormatAsYouTypeReplaceHyperlinks](http://msdn.microsoft.com/library/application-autoformatasyoutypereplacehyperlinks-property-excel%28Office.15%29.aspx)|
|[AutomationSecurity](http://msdn.microsoft.com/library/application-automationsecurity-property-excel%28Office.15%29.aspx)|
|[AutoPercentEntry](http://msdn.microsoft.com/library/application-autopercententry-property-excel%28Office.15%29.aspx)|
|[AutoRecover](http://msdn.microsoft.com/library/application-autorecover-property-excel%28Office.15%29.aspx)|
|[Build](http://msdn.microsoft.com/library/application-build-property-excel%28Office.15%29.aspx)|
|[CalculateBeforeSave](http://msdn.microsoft.com/library/application-calculatebeforesave-property-excel%28Office.15%29.aspx)|
|[Calculation](http://msdn.microsoft.com/library/application-calculation-property-excel%28Office.15%29.aspx)|
|[CalculationInterruptKey](http://msdn.microsoft.com/library/application-calculationinterruptkey-property-excel%28Office.15%29.aspx)|
|[CalculationState](http://msdn.microsoft.com/library/application-calculationstate-property-excel%28Office.15%29.aspx)|
|[CalculationVersion](http://msdn.microsoft.com/library/application-calculationversion-property-excel%28Office.15%29.aspx)|
|[Caller](http://msdn.microsoft.com/library/application-caller-property-excel%28Office.15%29.aspx)|
|[CanPlaySounds](http://msdn.microsoft.com/library/application-canplaysounds-property-excel%28Office.15%29.aspx)|
|[CanRecordSounds](http://msdn.microsoft.com/library/application-canrecordsounds-property-excel%28Office.15%29.aspx)|
|[Caption](http://msdn.microsoft.com/library/application-caption-property-excel%28Office.15%29.aspx)|
|[CellDragAndDrop](http://msdn.microsoft.com/library/application-celldraganddrop-property-excel%28Office.15%29.aspx)|
|[Cells](http://msdn.microsoft.com/library/application-cells-property-excel%28Office.15%29.aspx)|
|[ChartDataPointTrack](http://msdn.microsoft.com/library/application-chartdatapointtrack-property-excel%28Office.15%29.aspx)|
|[Charts](http://msdn.microsoft.com/library/application-charts-property-excel%28Office.15%29.aspx)|
|[ClipboardFormats](http://msdn.microsoft.com/library/application-clipboardformats-property-excel%28Office.15%29.aspx)|
|[ClusterConnector](http://msdn.microsoft.com/library/application-clusterconnector-property-excel%28Office.15%29.aspx)|
|[Columns](http://msdn.microsoft.com/library/application-columns-property-excel%28Office.15%29.aspx)|
|[COMAddIns](http://msdn.microsoft.com/library/application-comaddins-property-excel%28Office.15%29.aspx)|
|[CommandBars](http://msdn.microsoft.com/library/application-commandbars-property-excel%28Office.15%29.aspx)|
|[CommandUnderlines](http://msdn.microsoft.com/library/application-commandunderlines-property-excel%28Office.15%29.aspx)|
|[ConstrainNumeric](http://msdn.microsoft.com/library/application-constrainnumeric-property-excel%28Office.15%29.aspx)|
|[ControlCharacters](http://msdn.microsoft.com/library/application-controlcharacters-property-excel%28Office.15%29.aspx)|
|[CopyObjectsWithCells](http://msdn.microsoft.com/library/application-copyobjectswithcells-property-excel%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/application-creator-property-excel%28Office.15%29.aspx)|
|[Cursor](http://msdn.microsoft.com/library/application-cursor-property-excel%28Office.15%29.aspx)|
|[CursorMovement](http://msdn.microsoft.com/library/application-cursormovement-property-excel%28Office.15%29.aspx)|
|[CustomListCount](http://msdn.microsoft.com/library/application-customlistcount-property-excel%28Office.15%29.aspx)|
|[CutCopyMode](http://msdn.microsoft.com/library/application-cutcopymode-property-excel%28Office.15%29.aspx)|
|[DataEntryMode](http://msdn.microsoft.com/library/application-dataentrymode-property-excel%28Office.15%29.aspx)|
|[DDEAppReturnCode](http://msdn.microsoft.com/library/application-ddeappreturncode-property-excel%28Office.15%29.aspx)|
|[DecimalSeparator](http://msdn.microsoft.com/library/application-decimalseparator-property-excel%28Office.15%29.aspx)|
|[DefaultFilePath](http://msdn.microsoft.com/library/application-defaultfilepath-property-excel%28Office.15%29.aspx)|
|[DefaultSaveFormat](http://msdn.microsoft.com/library/application-defaultsaveformat-property-excel%28Office.15%29.aspx)|
|[DefaultSheetDirection](http://msdn.microsoft.com/library/application-defaultsheetdirection-property-excel%28Office.15%29.aspx)|
|[DefaultWebOptions](http://msdn.microsoft.com/library/application-defaultweboptions-property-excel%28Office.15%29.aspx)|
|[DeferAsyncQueries](http://msdn.microsoft.com/library/application-deferasyncqueries-property-excel%28Office.15%29.aspx)|
|[Dialogs](http://msdn.microsoft.com/library/application-dialogs-property-excel%28Office.15%29.aspx)|
|[DisplayAlerts](http://msdn.microsoft.com/library/application-displayalerts-property-excel%28Office.15%29.aspx)|
|[DisplayClipboardWindow](http://msdn.microsoft.com/library/application-displayclipboardwindow-property-excel%28Office.15%29.aspx)|
|[DisplayCommentIndicator](http://msdn.microsoft.com/library/application-displaycommentindicator-property-excel%28Office.15%29.aspx)|
|[DisplayDocumentActionTaskPane](http://msdn.microsoft.com/library/application-displaydocumentactiontaskpane-property-excel%28Office.15%29.aspx)|
|[DisplayDocumentInformationPanel](http://msdn.microsoft.com/library/application-displaydocumentinformationpanel-property-excel%28Office.15%29.aspx)|
|[DisplayExcel4Menus](http://msdn.microsoft.com/library/application-displayexcel4menus-property-excel%28Office.15%29.aspx)|
|[DisplayFormulaAutoComplete](http://msdn.microsoft.com/library/application-displayformulaautocomplete-property-excel%28Office.15%29.aspx)|
|[DisplayFormulaBar](http://msdn.microsoft.com/library/application-displayformulabar-property-excel%28Office.15%29.aspx)|
|[DisplayFullScreen](http://msdn.microsoft.com/library/application-displayfullscreen-property-excel%28Office.15%29.aspx)|
|[DisplayFunctionToolTips](http://msdn.microsoft.com/library/application-displayfunctiontooltips-property-excel%28Office.15%29.aspx)|
|[DisplayInsertOptions](http://msdn.microsoft.com/library/application-displayinsertoptions-property-excel%28Office.15%29.aspx)|
|[DisplayNoteIndicator](http://msdn.microsoft.com/library/application-displaynoteindicator-property-excel%28Office.15%29.aspx)|
|[DisplayPasteOptions](http://msdn.microsoft.com/library/application-displaypasteoptions-property-excel%28Office.15%29.aspx)|
|[DisplayRecentFiles](http://msdn.microsoft.com/library/application-displayrecentfiles-property-excel%28Office.15%29.aspx)|
|[DisplayScrollBars](http://msdn.microsoft.com/library/application-displayscrollbars-property-excel%28Office.15%29.aspx)|
|[DisplayStatusBar](http://msdn.microsoft.com/library/application-displaystatusbar-property-excel%28Office.15%29.aspx)|
|[EditDirectlyInCell](http://msdn.microsoft.com/library/application-editdirectlyincell-property-excel%28Office.15%29.aspx)|
|[EnableAutoComplete](http://msdn.microsoft.com/library/application-enableautocomplete-property-excel%28Office.15%29.aspx)|
|[EnableCancelKey](http://msdn.microsoft.com/library/application-enablecancelkey-property-excel%28Office.15%29.aspx)|
|[EnableCheckFileExtensions](http://msdn.microsoft.com/library/application-enablecheckfileextensions-property-excel%28Office.15%29.aspx)|
|[EnableEvents](http://msdn.microsoft.com/library/application-enableevents-property-excel%28Office.15%29.aspx)|
|[EnableLargeOperationAlert](http://msdn.microsoft.com/library/application-enablelargeoperationalert-property-excel%28Office.15%29.aspx)|
|[EnableLivePreview](http://msdn.microsoft.com/library/application-enablelivepreview-property-excel%28Office.15%29.aspx)|
|[EnableMacroAnimations](http://msdn.microsoft.com/library/application-enablemacroanimations-property-excel%28Office.15%29.aspx)|
|[EnableSound](http://msdn.microsoft.com/library/application-enablesound-property-excel%28Office.15%29.aspx)|
|[ErrorCheckingOptions](http://msdn.microsoft.com/library/application-errorcheckingoptions-property-excel%28Office.15%29.aspx)|
|[Excel4IntlMacroSheets](http://msdn.microsoft.com/library/application-excel4intlmacrosheets-property-excel%28Office.15%29.aspx)|
|[Excel4MacroSheets](http://msdn.microsoft.com/library/application-excel4macrosheets-property-excel%28Office.15%29.aspx)|
|[ExtendList](http://msdn.microsoft.com/library/application-extendlist-property-excel%28Office.15%29.aspx)|
|[FeatureInstall](http://msdn.microsoft.com/library/application-featureinstall-property-excel%28Office.15%29.aspx)|
|[FileConverters](http://msdn.microsoft.com/library/application-fileconverters-property-excel%28Office.15%29.aspx)|
|[FileDialog](http://msdn.microsoft.com/library/application-filedialog-property-excel%28Office.15%29.aspx)|
|[FileExportConverters](http://msdn.microsoft.com/library/application-fileexportconverters-property-excel%28Office.15%29.aspx)|
|[FileValidation](http://msdn.microsoft.com/library/application-filevalidation-property-excel%28Office.15%29.aspx)|
|[FileValidationPivot](http://msdn.microsoft.com/library/application-filevalidationpivot-property-excel%28Office.15%29.aspx)|
|[FindFormat](http://msdn.microsoft.com/library/application-findformat-property-excel%28Office.15%29.aspx)|
|[FixedDecimal](http://msdn.microsoft.com/library/application-fixeddecimal-property-excel%28Office.15%29.aspx)|
|[FixedDecimalPlaces](http://msdn.microsoft.com/library/application-fixeddecimalplaces-property-excel%28Office.15%29.aspx)|
|[FlashFill](http://msdn.microsoft.com/library/application-flashfill-property-excel%28Office.15%29.aspx)|
|[FlashFillMode](http://msdn.microsoft.com/library/application-flashfillmode-property-excel%28Office.15%29.aspx)|
|[FormulaBarHeight](http://msdn.microsoft.com/library/application-formulabarheight-property-excel%28Office.15%29.aspx)|
|[GenerateGetPivotData](http://msdn.microsoft.com/library/application-generategetpivotdata-property-excel%28Office.15%29.aspx)|
|[GenerateTableRefs](http://msdn.microsoft.com/library/application-generatetablerefs-property-excel%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/application-height-property-excel%28Office.15%29.aspx)|
|[HighQualityModeForGraphics](http://msdn.microsoft.com/library/application-highqualitymodeforgraphics-property-excel%28Office.15%29.aspx)|
|[Hinstance](http://msdn.microsoft.com/library/application-hinstance-property-excel%28Office.15%29.aspx)|
|[HinstancePtr](http://msdn.microsoft.com/library/application-hinstanceptr-property-excel%28Office.15%29.aspx)|
|[Hwnd](http://msdn.microsoft.com/library/application-hwnd-property-excel%28Office.15%29.aspx)|
|[IgnoreRemoteRequests](http://msdn.microsoft.com/library/application-ignoreremoterequests-property-excel%28Office.15%29.aspx)|
|[Interactive](http://msdn.microsoft.com/library/application-interactive-property-excel%28Office.15%29.aspx)|
|[International](http://msdn.microsoft.com/library/application-international-property-excel%28Office.15%29.aspx)|
|[IsSandboxed](http://msdn.microsoft.com/library/application-issandboxed-property-excel%28Office.15%29.aspx)|
|[Iteration](http://msdn.microsoft.com/library/application-iteration-property-excel%28Office.15%29.aspx)|
|[LanguageSettings](http://msdn.microsoft.com/library/application-languagesettings-property-excel%28Office.15%29.aspx)|
|[LargeOperationCellThousandCount](http://msdn.microsoft.com/library/application-largeoperationcellthousandcount-property-excel%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/application-left-property-excel%28Office.15%29.aspx)|
|[LibraryPath](http://msdn.microsoft.com/library/application-librarypath-property-excel%28Office.15%29.aspx)|
|[MailSession](http://msdn.microsoft.com/library/application-mailsession-property-excel%28Office.15%29.aspx)|
|[MailSystem](http://msdn.microsoft.com/library/application-mailsystem-property-excel%28Office.15%29.aspx)|
|[MapPaperSize](http://msdn.microsoft.com/library/application-mappapersize-property-excel%28Office.15%29.aspx)|
|[MathCoprocessorAvailable](http://msdn.microsoft.com/library/application-mathcoprocessoravailable-property-excel%28Office.15%29.aspx)|
|[MaxChange](http://msdn.microsoft.com/library/application-maxchange-property-excel%28Office.15%29.aspx)|
|[MaxIterations](http://msdn.microsoft.com/library/application-maxiterations-property-excel%28Office.15%29.aspx)|
|[MeasurementUnit](http://msdn.microsoft.com/library/application-measurementunit-property-excel%28Office.15%29.aspx)|
|[MergeInstances](http://msdn.microsoft.com/library/application-mergeinstances-property-excel%28Office.15%29.aspx)|
|[MouseAvailable](http://msdn.microsoft.com/library/application-mouseavailable-property-excel%28Office.15%29.aspx)|
|[MoveAfterReturn](http://msdn.microsoft.com/library/application-moveafterreturn-property-excel%28Office.15%29.aspx)|
|[MoveAfterReturnDirection](http://msdn.microsoft.com/library/application-moveafterreturndirection-property-excel%28Office.15%29.aspx)|
|[MultiThreadedCalculation](http://msdn.microsoft.com/library/application-multithreadedcalculation-property-excel%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/application-name-property-excel%28Office.15%29.aspx)|
|[Names](http://msdn.microsoft.com/library/application-names-property-excel%28Office.15%29.aspx)|
|[NetworkTemplatesPath](http://msdn.microsoft.com/library/application-networktemplatespath-property-excel%28Office.15%29.aspx)|
|[NewWorkbook](http://msdn.microsoft.com/library/application-newworkbook-property-excel%28Office.15%29.aspx)|
|[ODBCErrors](http://msdn.microsoft.com/library/application-odbcerrors-property-excel%28Office.15%29.aspx)|
|[ODBCTimeout](http://msdn.microsoft.com/library/application-odbctimeout-property-excel%28Office.15%29.aspx)|
|[OLEDBErrors](http://msdn.microsoft.com/library/application-oledberrors-property-excel%28Office.15%29.aspx)|
|[OnWindow](http://msdn.microsoft.com/library/application-onwindow-property-excel%28Office.15%29.aspx)|
|[OperatingSystem](http://msdn.microsoft.com/library/application-operatingsystem-property-excel%28Office.15%29.aspx)|
|[OrganizationName](http://msdn.microsoft.com/library/application-organizationname-property-excel%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/application-parent-property-excel%28Office.15%29.aspx)|
|[Path](http://msdn.microsoft.com/library/application-path-property-excel%28Office.15%29.aspx)|
|[PathSeparator](http://msdn.microsoft.com/library/application-pathseparator-property-excel%28Office.15%29.aspx)|
|[PivotTableSelection](http://msdn.microsoft.com/library/application-pivottableselection-property-excel%28Office.15%29.aspx)|
|[PreviousSelections](http://msdn.microsoft.com/library/application-previousselections-property-excel%28Office.15%29.aspx)|
|[PrintCommunication](http://msdn.microsoft.com/library/application-printcommunication-property-excel%28Office.15%29.aspx)|
|[ProductCode](http://msdn.microsoft.com/library/application-productcode-property-excel%28Office.15%29.aspx)|
|[PromptForSummaryInfo](http://msdn.microsoft.com/library/application-promptforsummaryinfo-property-excel%28Office.15%29.aspx)|
|[ProtectedViewWindows](http://msdn.microsoft.com/library/application-protectedviewwindows-property-excel%28Office.15%29.aspx)|
|[QuickAnalysis](http://msdn.microsoft.com/library/application-quickanalysis-property-excel%28Office.15%29.aspx)|
|[Range](http://msdn.microsoft.com/library/application-range-property-excel%28Office.15%29.aspx)|
|[Ready](http://msdn.microsoft.com/library/application-ready-property-excel%28Office.15%29.aspx)|
|[RecentFiles](http://msdn.microsoft.com/library/application-recentfiles-property-excel%28Office.15%29.aspx)|
|[RecordRelative](http://msdn.microsoft.com/library/application-recordrelative-property-excel%28Office.15%29.aspx)|
|[ReferenceStyle](http://msdn.microsoft.com/library/application-referencestyle-property-excel%28Office.15%29.aspx)|
|[RegisteredFunctions](http://msdn.microsoft.com/library/application-registeredfunctions-property-excel%28Office.15%29.aspx)|
|[ReplaceFormat](http://msdn.microsoft.com/library/application-replaceformat-property-excel%28Office.15%29.aspx)|
|[RollZoom](http://msdn.microsoft.com/library/application-rollzoom-property-excel%28Office.15%29.aspx)|
|[Rows](http://msdn.microsoft.com/library/application-rows-property-excel%28Office.15%29.aspx)|
|[RTD](http://msdn.microsoft.com/library/application-rtd-property-excel%28Office.15%29.aspx)|
|[ScreenUpdating](http://msdn.microsoft.com/library/application-screenupdating-property-excel%28Office.15%29.aspx)|
|[Selection](http://msdn.microsoft.com/library/application-selection-property-excel%28Office.15%29.aspx)|
|[Sheets](http://msdn.microsoft.com/library/application-sheets-property-excel%28Office.15%29.aspx)|
|[SheetsInNewWorkbook](http://msdn.microsoft.com/library/application-sheetsinnewworkbook-property-excel%28Office.15%29.aspx)|
|[ShowChartTipNames](http://msdn.microsoft.com/library/application-showcharttipnames-property-excel%28Office.15%29.aspx)|
|[ShowChartTipValues](http://msdn.microsoft.com/library/application-showcharttipvalues-property-excel%28Office.15%29.aspx)|
|[ShowDevTools](http://msdn.microsoft.com/library/application-showdevtools-property-excel%28Office.15%29.aspx)|
|[ShowMenuFloaties](http://msdn.microsoft.com/library/application-showmenufloaties-property-excel%28Office.15%29.aspx)|
|[ShowQuickAnalysis](http://msdn.microsoft.com/library/application-showquickanalysis-property-excel%28Office.15%29.aspx)|
|[ShowSelectionFloaties](http://msdn.microsoft.com/library/application-showselectionfloaties-property-excel%28Office.15%29.aspx)|
|[ShowStartupDialog](http://msdn.microsoft.com/library/application-showstartupdialog-property-excel%28Office.15%29.aspx)|
|[ShowToolTips](http://msdn.microsoft.com/library/application-showtooltips-property-excel%28Office.15%29.aspx)|
|[SmartArtColors](http://msdn.microsoft.com/library/application-smartartcolors-property-excel%28Office.15%29.aspx)|
|[SmartArtLayouts](http://msdn.microsoft.com/library/application-smartartlayouts-property-excel%28Office.15%29.aspx)|
|[SmartArtQuickStyles](http://msdn.microsoft.com/library/application-smartartquickstyles-property-excel%28Office.15%29.aspx)|
|[Speech](http://msdn.microsoft.com/library/application-speech-property-excel%28Office.15%29.aspx)|
|[SpellingOptions](http://msdn.microsoft.com/library/application-spellingoptions-property-excel%28Office.15%29.aspx)|
|[StandardFont](http://msdn.microsoft.com/library/application-standardfont-property-excel%28Office.15%29.aspx)|
|[StandardFontSize](http://msdn.microsoft.com/library/application-standardfontsize-property-excel%28Office.15%29.aspx)|
|[StartupPath](http://msdn.microsoft.com/library/application-startuppath-property-excel%28Office.15%29.aspx)|
|[StatusBar](http://msdn.microsoft.com/library/application-statusbar-property-excel%28Office.15%29.aspx)|
|[TemplatesPath](http://msdn.microsoft.com/library/application-templatespath-property-excel%28Office.15%29.aspx)|
|[ThisCell](http://msdn.microsoft.com/library/application-thiscell-property-excel%28Office.15%29.aspx)|
|[ThisWorkbook](http://msdn.microsoft.com/library/application-thisworkbook-property-excel%28Office.15%29.aspx)|
|[ThousandsSeparator](http://msdn.microsoft.com/library/application-thousandsseparator-property-excel%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/application-top-property-excel%28Office.15%29.aspx)|
|[TransitionMenuKey](http://msdn.microsoft.com/library/application-transitionmenukey-property-excel%28Office.15%29.aspx)|
|[TransitionMenuKeyAction](http://msdn.microsoft.com/library/application-transitionmenukeyaction-property-excel%28Office.15%29.aspx)|
|[TransitionNavigKeys](http://msdn.microsoft.com/library/application-transitionnavigkeys-property-excel%28Office.15%29.aspx)|
|[UsableHeight](http://msdn.microsoft.com/library/application-usableheight-property-excel%28Office.15%29.aspx)|
|[UsableWidth](http://msdn.microsoft.com/library/application-usablewidth-property-excel%28Office.15%29.aspx)|
|[UseClusterConnector](http://msdn.microsoft.com/library/application-useclusterconnector-property-excel%28Office.15%29.aspx)|
|[UsedObjects](http://msdn.microsoft.com/library/application-usedobjects-property-excel%28Office.15%29.aspx)|
|[UserControl](http://msdn.microsoft.com/library/application-usercontrol-property-excel%28Office.15%29.aspx)|
|[UserLibraryPath](http://msdn.microsoft.com/library/application-userlibrarypath-property-excel%28Office.15%29.aspx)|
|[UserName](http://msdn.microsoft.com/library/application-username-property-excel%28Office.15%29.aspx)|
|[UseSystemSeparators](http://msdn.microsoft.com/library/application-usesystemseparators-property-excel%28Office.15%29.aspx)|
|[Value](http://msdn.microsoft.com/library/application-value-property-excel%28Office.15%29.aspx)|
|[VBE](http://msdn.microsoft.com/library/application-vbe-property-excel%28Office.15%29.aspx)|
|[Version](http://msdn.microsoft.com/library/application-version-property-excel%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/application-visible-property-excel%28Office.15%29.aspx)|
|[WarnOnFunctionNameConflict](http://msdn.microsoft.com/library/application-warnonfunctionnameconflict-property-excel%28Office.15%29.aspx)|
|[Watches](http://msdn.microsoft.com/library/application-watches-property-excel%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/application-width-property-excel%28Office.15%29.aspx)|
|[Windows](http://msdn.microsoft.com/library/application-windows-property-excel%28Office.15%29.aspx)|
|[WindowsForPens](http://msdn.microsoft.com/library/application-windowsforpens-property-excel%28Office.15%29.aspx)|
|[WindowState](http://msdn.microsoft.com/library/application-windowstate-property-excel%28Office.15%29.aspx)|
|[Workbooks](http://msdn.microsoft.com/library/application-workbooks-property-excel%28Office.15%29.aspx)|
|[WorksheetFunction](http://msdn.microsoft.com/library/application-worksheetfunction-property-excel%28Office.15%29.aspx)|
|[Worksheets](http://msdn.microsoft.com/library/application-worksheets-property-excel%28Office.15%29.aspx)|
|[EnableAnimations](http://msdn.microsoft.com/library/application-enableanimations-property-excel%28Office.15%29.aspx)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/object-model-excel-vba-reference%28Office.15%29.aspx)
<<<<<<< HEAD
=======

>>>>>>> d7667e83d23dbf8ebf5bf068ba6fed14c840c0f5

