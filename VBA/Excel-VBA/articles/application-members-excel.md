---
title: Application Members (Excel)
ms.prod: EXCEL
ms.assetid: 4cb9ca42-8d07-cc9c-2d80-4eb9a5921e1e
---


# Application Members (Excel)
Represents the entire Microsoft Excel application.

Represents the entire Microsoft Excel application.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[AfterCalculate](application-aftercalculate-event-excel.md)|The  **AfterCalculate** event occurs when all pending refresh activity (both synchronous and asynchronous) and all of the resultant calculation activities have been completed.|
|[NewWorkbook](application-newworkbook-event-excel.md)|Occurs when a new workbook is created.|
|[ProtectedViewWindowActivate](application-protectedviewwindowactivate-event-excel.md)|Occurs when a  **Protected View** window is activated.|
|[ProtectedViewWindowBeforeClose](application-protectedviewwindowbeforeclose-event-excel.md)|Occurs immediately before a  **Protected View** window or a workbook in a **Protected View** window closes.|
|[ProtectedViewWindowBeforeEdit](application-protectedviewwindowbeforeedit-event-excel.md)|Occurs immediately before editing is enabled on the workbook in the specified  **Protected View** window.|
|[ProtectedViewWindowDeactivate](application-protectedviewwindowdeactivate-event-excel.md)|Occurs when a  **Protected View** window is deactivated.|
|[ProtectedViewWindowOpen](application-protectedviewwindowopen-event-excel.md)|Occurs when a workbook is opened in a  **Protected View** window.|
|[ProtectedViewWindowResize](application-protectedviewwindowresize-event-excel.md)|Occurs when any  **Protected View** window is resized.|
|[SheetActivate](application-sheetactivate-event-excel.md)|Occurs when any sheet is activated.|
|[SheetBeforeDelete](application-sheetbeforedelete-event-excel.md)||
|[SheetBeforeDoubleClick](application-sheetbeforedoubleclick-event-excel.md)|Occurs when any worksheet is double-clicked, before the default double-click action.|
|[SheetBeforeRightClick](application-sheetbeforerightclick-event-excel.md)|Occurs when any worksheet is right-clicked, before the default right-click action.|
|[SheetCalculate](application-sheetcalculate-event-excel.md)|Occurs after any worksheet is recalculated or after any changed data is plotted on a chart.|
|[SheetChange](application-sheetchange-event-excel.md)|Occurs when cells in any worksheet are changed by the user or by an external link.|
|[SheetDeactivate](application-sheetdeactivate-event-excel.md)|Occurs when any sheet is deactivated.|
|[SheetFollowHyperlink](application-sheetfollowhyperlink-event-excel.md)|Occurs when you click any hyperlink in Microsoft Excel. For worksheet-level events, see the Help topic for the  **[FollowHyperlink](worksheet-followhyperlink-event-excel.md)** event.|
|[SheetLensGalleryRenderComplete](application-sheetlensgalleryrendercomplete-event-excel.md)|Occurs after a callout gallery's icons (dynamic &; static) have finished rendering.|
|[SheetPivotTableAfterValueChange](application-sheetpivottableaftervaluechange-event-excel.md)|Occurs after a cell or range of cells inside a PivotTable are edited or recalculated (for cells that contain formulas).|
|[SheetPivotTableBeforeAllocateChanges](application-sheetpivottablebeforeallocatechanges-event-excel.md)|Occurs before changes are applied to a PivotTable.|
|[SheetPivotTableBeforeCommitChanges](application-sheetpivottablebeforecommitchanges-event-excel.md)|Occurs before changes are committed against the OLAP data source for a PivotTable.|
|[SheetPivotTableBeforeDiscardChanges](application-sheetpivottablebeforediscardchanges-event-excel.md)|Occurs before changes to a PivotTable are discarded.|
|[SheetPivotTableUpdate](application-sheetpivottableupdate-event-excel.md)|Occurs after the sheet of the PivotTable report has been updated.|
|[SheetSelectionChange](application-sheetselectionchange-event-excel.md)|Occurs when the selection changes on any worksheet (doesn't occur if the selection is on a chart sheet).|
|[SheetTableUpdate](application-sheettableupdate-event-excel.md)|Occurs when a table on a worksheet is updated.|
|[WindowActivate](application-windowactivate-event-excel.md)|Occurs when any workbook window is activated.|
|[WindowDeactivate](application-windowdeactivate-event-excel.md)|Occurs when any workbook window is deactivated.|
|[WindowResize](application-windowresize-event-excel.md)|Occurs when any workbook window is resized.|
|[WorkbookActivate](application-workbookactivate-event-excel.md)|Occurs when any workbook is activated.|
|[WorkbookAddinInstall](application-workbookaddininstall-event-excel.md)|Occurs when a workbook is installed as an add-in.|
|[WorkbookAddinUninstall](application-workbookaddinuninstall-event-excel.md)|Occurs when any add-in workbook is uninstalled.|
|[WorkbookAfterSave](application-workbookaftersave-event-excel.md)|Occurs after the workbook is saved.|
|[WorkbookAfterXmlExport](application-workbookafterxmlexport-event-excel.md)|Occurs after Microsoft Excel saves or exports XML data from the specified workbook.|
|[WorkbookAfterXmlImport](application-workbookafterxmlimport-event-excel.md)|Occurs after an existing XML data connection is refreshed, or new XML data is imported into any open Microsoft Excel workbook.|
|[WorkbookBeforeClose](application-workbookbeforeclose-event-excel.md)|Occurs immediately before any open workbook closes.|
|[WorkbookBeforePrint](application-workbookbeforeprint-event-excel.md)|Occurs before any open workbook is printed.|
|[WorkbookBeforeSave](application-workbookbeforesave-event-excel.md)|Occurs before any open workbook is saved.|
|[WorkbookBeforeXmlExport](application-workbookbeforexmlexport-event-excel.md)|Occurs before Microsoft Excel saves or exports XML data from the specified workbook.|
|[WorkbookBeforeXmlImport](application-workbookbeforexmlimport-event-excel.md)|Occurs before an existing XML data connection is refreshed, or new XML data is imported into any open Microsoft Excel workbook.|
|[WorkbookDeactivate](application-workbookdeactivate-event-excel.md)|Occurs when any open workbook is deactivated.|
|[WorkbookModelChange](application-workbookmodelchange-event-excel.md)|Occurs when the data model is updated.|
|[WorkbookNewChart](application-workbooknewchart-event-excel.md)|Occurs when a new chart is created in any open workbook.|
|[WorkbookNewSheet](application-workbooknewsheet-event-excel.md)|Occurs when a new sheet is created in any open workbook.|
|[WorkbookOpen](application-workbookopen-event-excel.md)|Occurs when a workbook is opened.|
|[WorkbookPivotTableCloseConnection](application-workbookpivottablecloseconnection-event-excel.md)|Occurs after a PivotTable report connection has been closed.|
|[WorkbookPivotTableOpenConnection](application-workbookpivottableopenconnection-event-excel.md)|Occurs after a PivotTable report connection has been opened.|
|[WorkbookRowsetComplete](application-workbookrowsetcomplete-event-excel.md)|The  **WorkbookRowsetComplete** event occurs when the user either drills through the recordset or invokes the rowset action on an OLAP PivotTable.|
|[WorkbookSync](application-workbooksync-event-excel.md)|This object or member has been deprecated, but it remains part of the object model for backward compatibility. You should not use it in new applications.|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[ActivateMicrosoftApp](application-activatemicrosoftapp-method-excel.md)|Activates a Microsoft application. If the application is already running, this method activates the running application. If the application isn't running, this method starts a new instance of the application.|
|[AddCustomList](application-addcustomlist-method-excel.md)|Adds a custom list for custom autofill and/or custom sort.|
|[Calculate](application-calculate-method-excel.md)|Calculates all open workbooks, a specific worksheet in a workbook, or a specified range of cells on a worksheet, as shown in the following table.|
|[CalculateFull](application-calculatefull-method-excel.md)|Forces a full calculation of the data in all open workbooks.|
|[CalculateFullRebuild](application-calculatefullrebuild-method-excel.md)|For all open workbooks, forces a full calculation of the data and rebuilds the dependencies.|
|[CalculateUntilAsyncQueriesDone](application-calculateuntilasyncqueriesdone-method-excel.md)|Runs all pending queries to OLEDB and OLAP data sources.|
|[CentimetersToPoints](application-centimeterstopoints-method-excel.md)|Converts a measurement from centimeters to points (one point equals 0.035 centimeters).|
|[CheckAbort](application-checkabort-method-excel.md)|Stops recalculation in a Microsoft Excel application.|
|[CheckSpelling](application-checkspelling-method-excel.md)|Checks the spelling of a single word.|
|[ConvertFormula](application-convertformula-method-excel.md)|Converts cell references in a formula between the A1 and R1C1 reference styles, between relative and absolute references, or both.  **Variant** .|
|[DDEExecute](application-ddeexecute-method-excel.md)|Runs a command or performs some other action or actions in another application by way of the specified DDE channel.|
|[DDEInitiate](application-ddeinitiate-method-excel.md)|Opens a DDE channel to an application.|
|[DDEPoke](application-ddepoke-method-excel.md)|Sends data to an application.|
|[DDERequest](application-dderequest-method-excel.md)|Requests information from the specified application. This method always returns an array.|
|[DDETerminate](application-ddeterminate-method-excel.md)|Closes a channel to another application.|
|[DeleteCustomList](application-deletecustomlist-method-excel.md)|Deletes a custom list.|
|[DisplayXMLSourcePane](application-displayxmlsourcepane-method-excel.md)|Opens the  **XML Source** task pane and displays the XML map specified by the _XmlMap_ argument.|
|[DoubleClick](application-doubleclick-method-excel.md)|Equivalent to double-clicking the active cell.|
|[Evaluate](application-evaluate-method-excel.md)|Converts a Microsoft Excel name to an object or a value.|
|[ExecuteExcel4Macro](application-executeexcel4macro-method-excel.md)|Runs a Microsoft Excel 4.0 macro function and then returns the result of the function. The return type depends on the function.|
|[FindFile](application-findfile-method-excel.md)|Displays the  **Open** dialog box.|
|[GetCustomListContents](application-getcustomlistcontents-method-excel.md)|Returns a custom list (an array of strings).|
|[GetCustomListNum](application-getcustomlistnum-method-excel.md)|Returns the custom list number for an array of strings. You can use this method to match both built-in lists and custom-defined lists.|
|[GetOpenFilename](application-getopenfilename-method-excel.md)|Displays the standard  **Open** dialog box and gets a file name from the user without actually opening any files.|
|[GetPhonetic](application-getphonetic-method-excel.md)|Returns the Japanese phonetic text of the specified text string. This method is available to you only if you have selected or installed Japanese language support for Microsoft Office.|
|[GetSaveAsFilename](application-getsaveasfilename-method-excel.md)|Displays the standard  **Save As** dialog box and gets a file name from the user without actually saving any files.|
|[Goto](application-goto-method-excel.md)|Selects any range or Visual Basic procedure in any workbook, and activates that workbook if it's not already active.|
|[Help](application-help-method-excel.md)|Displays a Help topic.|
|[InchesToPoints](application-inchestopoints-method-excel.md)|Converts a measurement from inches to points.|
|[InputBox](application-inputbox-method-excel.md)|Displays a dialog box for user input. Returns the information entered in the dialog box.|
|[Intersect](application-intersect-method-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents the rectangular intersection of two or more ranges.|
|[MacroOptions](application-macrooptions-method-excel.md)|Corresponds to options in the  **Macro Options** dialog box. You can also use this method to display a user defined function (UDF) in a built-in or new category within the **Insert Function** dialog box.|
|[MailLogoff](application-maillogoff-method-excel.md)|Closes a MAPI mail session established by Microsoft Excel.|
|[MailLogon](application-maillogon-method-excel.md)|Logs in to MAPI Mail or Microsoft Exchange and establishes a mail session. If Microsoft Mail isn't already running, you must use this method to establish a mail session before mail or document routing functions can be used.|
|[NextLetter](application-nextletter-method-excel.md)|You have requested Help for a Visual Basic keyword used only on the Macintosh. For information about this keyword, consult the language reference Help included with Microsoft Office Macintosh Edition.|
|[OnKey](application-onkey-method-excel.md)|Runs a specified procedure when a particular key or key combination is pressed.|
|[OnRepeat](application-onrepeat-method-excel.md)|Sets the  **Repeat** item and the name of the procedure that will run if you choose the **Repeat** command after running the procedure that sets this property.|
|[OnTime](application-ontime-method-excel.md)|Schedules a procedure to be run at a specified time in the future (either at a specific time of day or after a specific amount of time has passed).|
|[OnUndo](application-onundo-method-excel.md)|Sets the text of the  **Undo** command and the name of the procedure that's run if you choose the **Undo** command after running the procedure that sets this property.|
|[Quit](application-quit-method-excel.md)|Quits Microsoft Excel.|
|[RecordMacro](application-recordmacro-method-excel.md)|Records code if the macro recorder is on.|
|[RegisterXLL](application-registerxll-method-excel.md)|Loads an XLL code resource and automatically registers the functions and commands contained in the resource.|
|[Repeat](application-repeat-method-excel.md)|Repeats the last user-interface action.|
|[Run](application-run-method-excel.md)|Runs a macro or calls a function. This can be used to run a macro written in Visual Basic or the Microsoft Excel macro language, or to run a function in a DLL or XLL.|
|[SendKeys](application-sendkeys-method-excel.md)|Sends keystrokes to the active application.|
|[SharePointVersion](application-sharepointversion-method-excel.md)|Returns the version number of SharePoint Foundation instances running at site for the specified URL.|
|[Undo](application-undo-method-excel.md)|Cancels the last user-interface action.|
|[Union](application-union-method-excel.md)|Returns the union of two or more ranges.|
|[Volatile](application-volatile-method-excel.md)|Marks a user-defined function as volatile. A volatile function must be recalculated whenever calculation occurs in any cells on the worksheet. A nonvolatile function is recalculated only when the input variables change. This method has no effect if it's not inside a user-defined function used to calculate a worksheet cell.|
|[Wait](application-wait-method-excel.md)|Pauses a running macro until a specified time. Returns  **True** if the specified time has arrived.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[ActiveCell](application-activecell-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents the active cell in the active window (the window on top) or in the specified window. If the window isn't displaying a worksheet, this property fails. Read-only.|
|[ActiveChart](application-activechart-property-excel.md)|Returns a  **[Chart](chart-object-excel.md)** object that represents the active chart (either an embedded chart or a chart sheet). An embedded chart is considered active when it's either selected or activated. When no chart is active, this property returns **Nothing** .|
|[ActiveEncryptionSession](application-activeencryptionsession-property-excel.md)|Returns a  **Long** that represents the encryption session associated with the active document. Read-only.|
|[ActivePrinter](application-activeprinter-property-excel.md)|Returns or sets the name of the active printer. Read/write  **String** .|
|[ActiveProtectedViewWindow](application-activeprotectedviewwindow-property-excel.md)|Returns a  **[ProtectedViewWindow](protectedviewwindow-object-excel.md)** object that represents the active **Protected View** window (the window on top). Read-only. Returns **Nothing** if there are no **Protected View** windows open. Read-only|
|[ActiveSheet](application-activesheet-property-excel.md)|Returns an object that represents the active sheet (the sheet on top) in the active workbook or in the specified window or workbook. Returns  **Nothing** if no sheet is active.|
|[ActiveWindow](application-activewindow-property-excel.md)|Returns a  **[Window](window-object-excel.md)** object that represents the active window (the window on top). Read-only. Returns **Nothing** if there are no windows open.|
|[ActiveWorkbook](application-activeworkbook-property-excel.md)|Returns a  **[Workbook](workbook-object-excel.md)** object that represents the workbook in the active window (the window on top). Read-only. Returns **Nothing** if there are no windows open or if either the Info window or the Clipboard window is the active window.|
|[AddIns](application-addins-property-excel.md)|Returns an  **[AddIns](addins-object-excel.md)** collection that represents all the add-ins listed in the **Add-Ins** dialog box ( **Add-Ins** command on the **Developer** tab). Read-only.|
|[AddIns2](application-addins2-property-excel.md)|Returns an  **[AddIns2](addins2-object-excel.md)** collection that represents all the add-ins that are currently available or open in Microsoft Excel, regardless of whether they are installed. Read-only|
|[AlertBeforeOverwriting](application-alertbeforeoverwriting-property-excel.md)| **True** if Microsoft Excel displays a message before overwriting nonblank cells during a drag-and-drop editing operation. Read/write **Boolean** .|
|[AltStartupPath](application-altstartuppath-property-excel.md)|Returns or sets the name of the alternate startup folder. Read/write  **String** .|
|[AlwaysUseClearType](application-alwaysusecleartype-property-excel.md)|Returns or sets a  **Boolean** that represents whether to use ClearType to display fonts in the menu, ribbon, and dialog box text. Read/write **Boolean** .|
|[Application](application-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[ArbitraryXMLSupportAvailable](application-arbitraryxmlsupportavailable-property-excel.md)| Returns a **Boolean** value that indicates whether the XML features in Microsoft Excel are available. Read-only.|
|[AskToUpdateLinks](application-asktoupdatelinks-property-excel.md)| **True** if Microsoft Excel asks the user to update links when opening files with links. **False** if links are automatically updated with no dialog box. Read/write **Boolean** .|
|[Assistance](application-assistance-property-excel.md)|Returns an  **[IAssistance](iassistance-object-office.md)** object for Microsoft Excel that represents the Microsoft Office Help Viewer. Read-only.|
|[AutoCorrect](application-autocorrect-property-excel.md)|Returns an  **[AutoCorrect](autocorrect-object-excel.md)** object that represents the Microsoft Excel AutoCorrect attributes. Read-only.|
|[AutoFormatAsYouTypeReplaceHyperlinks](application-autoformatasyoutypereplacehyperlinks-property-excel.md)| **True** (default) if Microsoft Excel automatically formats hyperlinks as you type. **False** if Excel does not automatically format hyperlinks as you type. Read/write **Boolean** .|
|[AutomationSecurity](application-automationsecurity-property-excel.md)|Returns or sets an  **[MsoAutomationSecurity](msoautomationsecurity-enumeration-office.md)** constant that represents the security mode Microsoft Excel uses when programmatically opening files. Read/write.|
|[AutoPercentEntry](application-autopercententry-property-excel.md)| **True** if entries in cells formatted as percentages aren't automatically multiplied by 100 as soon as they are entered. Read/write **Boolean** .|
|[AutoRecover](application-autorecover-property-excel.md)|Returns an  **[AutoRecover](autorecover-object-excel.md)** object, which backs up all file formats on a timed interval.|
|[Build](application-build-property-excel.md)|Returns the Microsoft Excel build number. Read-only  **Long** .|
|[CalculateBeforeSave](application-calculatebeforesave-property-excel.md)| **True** if workbooks are calculated before they're saved to disk (if the **[Calculation](application-calculation-property-excel.md)** property is set to **xlManual** ). This property is preserved even if you change the **Calculation** property. Read/write **Boolean** .|
|[Calculation](application-calculation-property-excel.md)|Returns or sets a  **[XlCalculation](xlcalculation-enumeration-excel.md)** value that represents the calculation mode.|
|[CalculationInterruptKey](application-calculationinterruptkey-property-excel.md)|Sets or returns an  **[XlCalculationInterruptKey](xlcalculationinterruptkey-enumeration-excel.md)** constant that specifies the key that can interrupt Microsoft Excel when performing calculations. Read/write.|
|[CalculationState](application-calculationstate-property-excel.md)|Returns an  **[XlCalculationState](xlcalculationstate-enumeration-excel.md)** constant that indicates the calculation state of the application, for any calculations that are being performed in Microsoft Excel. Read-only.|
|[CalculationVersion](application-calculationversion-property-excel.md)|Returns a number whose rightmost four digits are the minor calculation engine version number, and whose other digits (on the left) are the major version of Microsoft Excel. Read-only  **Long** .|
|[Caller](application-caller-property-excel.md)|Returns information about how Visual Basic was called (for more information, see the Remarks section).|
|[CanPlaySounds](application-canplaysounds-property-excel.md)|This property should not be used. Sound notes have been removed from Microsoft Excel.|
|[CanRecordSounds](application-canrecordsounds-property-excel.md)|This property should not be used. Sound notes have been removed from Microsoft Excel.|
|[Caption](application-caption-property-excel.md)|Returns or sets a  **String** value that represents the name that appears in the title bar of the main Microsoft Excel window.|
|[CellDragAndDrop](application-celldraganddrop-property-excel.md)| **True** if dragging and dropping cells is enabled. Read/write **Boolean** .|
|[Cells](application-cells-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents all the cells on the active worksheet. If the active document is not a worksheet, this property fails.|
|[ChartDataPointTrack](application-chartdatapointtrack-property-excel.md)| **True** will cause all charts in newly created documents to use the cell reference tracking behavior. **Boolean**|
|[Charts](application-charts-property-excel.md)|Returns a  **[Sheets](sheets-object-excel.md)** collection that represents all the chart sheets in the active workbook.|
|[ClipboardFormats](application-clipboardformats-property-excel.md)|Returns the formats that are currently on the Clipboard, as an array of numeric values. To determine whether a particular format is on the Clipboard, compare each element in the array with the appropriate constant listed in the Remarks section. Read-only  **Variant** .|
|[ClusterConnector](application-clusterconnector-property-excel.md)|Returns or sets the name of the High Performance Computing (HPC) Cluster Connector that is used to run user-defined functions in XLL add-ins. Read/write|
|[Columns](application-columns-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents all the columns on the active worksheet. If the active document isn't a worksheet, the **Columns** property fails.|
|[COMAddIns](application-comaddins-property-excel.md)|Returns the  **[COMAddIns](comaddins-object-office.md)** collection for Microsoft Excel, which represents the currently installed COM add-ins. Read-only.|
|[CommandBars](application-commandbars-property-excel.md)|Returns a  **[CommandBars](commandbars-object-office.md)** object that represents the Microsoft Excel command bars. Read-only.|
|[CommandUnderlines](application-commandunderlines-property-excel.md)|Returns or sets the state of the command underlines in Microsoft Excel for the Macintosh. Can be one of the constants of  **[XlCommandUnderlines](xlcommandunderlines-enumeration-excel.md)** . Read/write **Long** .|
|[ConstrainNumeric](application-constrainnumeric-property-excel.md)| **True** if handwriting recognition is limited to numbers and punctuation only. Read/write **Boolean** .|
|[ControlCharacters](application-controlcharacters-property-excel.md)| **True** if Microsoft Excel displays control characters for right-to-left languages. Read/write **Boolean** .|
|[CopyObjectsWithCells](application-copyobjectswithcells-property-excel.md)| **True** if objects are cut, copied, extracted, and sorted with cells. Read/write **Boolean** .|
|[Creator](application-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[Cursor](application-cursor-property-excel.md)|Returns or sets the appearance of the mouse pointer in Microsoft Excel. Read/write  **[XlMousePointer](xlmousepointer-enumeration-excel.md)** .|
|[CursorMovement](application-cursormovement-property-excel.md)|Returns or sets a value that indicates whether a visual cursor or a logical cursor is used. Can be one of the following constants:  **xlVisualCursor** or **xlLogicalCursor** . Read/write **Long** .|
|[CustomListCount](application-customlistcount-property-excel.md)|Returns the number of defined custom lists (including built-in lists). Read-only  **Long** .|
|[CutCopyMode](application-cutcopymode-property-excel.md)|Returns or sets the status of Cut or Copy mode. Can be  **True** , **False** , or an **[XLCutCopyMode](xlcutcopymode-enumeration-excel.md)** constant, as shown in the following tables. Read/write **Long** .|
|[DataEntryMode](application-dataentrymode-property-excel.md)|Returns or sets Data Entry mode, as shown in the following table. When in Data Entry mode, you can enter data only in the unlocked cells in the currently selected range. Read/write  **Long** .|
|[DDEAppReturnCode](application-ddeappreturncode-property-excel.md)|Returns the application-specific DDE return code that was contained in the last DDE acknowledge message received by Microsoft Excel. Read-only  **Long** .|
|[DecimalSeparator](application-decimalseparator-property-excel.md)|Sets or returns the character used for the decimal separator as a  **String** . Read/write.|
|[DefaultFilePath](application-defaultfilepath-property-excel.md)|Returns or sets the default path that Microsoft Excel uses when it opens files. Read/write  **String** .|
|[DefaultSaveFormat](application-defaultsaveformat-property-excel.md)|Returns or sets the default format for saving files. For a list of valid constants, see the  **[FileFormat](workbook-fileformat-property-excel.md)** property. Read/write **Long** .|
|[DefaultSheetDirection](application-defaultsheetdirection-property-excel.md)|Returns or sets the default direction in which Microsoft Excel displays new windows and worksheets. Can be one of the following constants:  **xlRTL** (right to left) or **xlLTR** (left to right). Read/write **Long** .|
|[DefaultWebOptions](application-defaultweboptions-property-excel.md)|Returns the  **[DefaultWebOptions](defaultweboptions-object-excel.md)** object that contains global application-level attributes used by Microsoft Excel whenever you save a document as a Web page or open a Web page. Read-only.|
|[DeferAsyncQueries](application-deferasyncqueries-property-excel.md)|Gets or sets whether asychronous queries to OLAP data sources are executed when a worksheet is calculated by VBA code. Read/write  **Boolean** .|
|[Dialogs](application-dialogs-property-excel.md)|Returns a  **[Dialogs](dialogs-object-excel.md)** collection that represents all built-in dialog boxes. Read-only.|
|[DisplayAlerts](application-displayalerts-property-excel.md)| **True** if Microsoft Excel displays certain alerts and messages while a macro is running. Read/write **Boolean** .|
|[DisplayClipboardWindow](application-displayclipboardwindow-property-excel.md)|Returns  **True** if the Microsoft Office Clipboard can be displayed. Read/write **Boolean** .|
|[DisplayCommentIndicator](application-displaycommentindicator-property-excel.md)|Returns or sets the way cells display comments and indicators. Can be one of the  **[XlCommentDisplayMode](xlcommentdisplaymode-enumeration-excel.md)** constants.|
|[DisplayDocumentActionTaskPane](application-displaydocumentactiontaskpane-property-excel.md)|Set to  **True** to display the **Document Actions** task pane; set to **False** to hide the **Document Actions** task pane. Read/write **Boolean** .|
|[DisplayDocumentInformationPanel](application-displaydocumentinformationpanel-property-excel.md)|Returns or sets a  **Boolean** that represents whether the document properties panel is displayed. Read/write **Boolean** .|
|[DisplayExcel4Menus](application-displayexcel4menus-property-excel.md)| **True** if Microsoft Excel displays version 4.0 menu bars. Read/write **Boolean** .|
|[DisplayFormulaAutoComplete](application-displayformulaautocomplete-property-excel.md)|Gets or sets whether to show a list of relevant functions and defined names when building cell formulas. Read/write  **Boolean** .|
|[DisplayFormulaBar](application-displayformulabar-property-excel.md)| **True** if the formula bar is displayed. Read/write **Boolean** .|
|[DisplayFullScreen](application-displayfullscreen-property-excel.md)| **True** if Microsoft Excel is in full-screen mode. Read/write **Boolean** .|
|[DisplayFunctionToolTips](application-displayfunctiontooltips-property-excel.md)| **True** if function ToolTips can be displayed. Read/write **Boolean** .|
|[DisplayInsertOptions](application-displayinsertoptions-property-excel.md)| **True** if the **Insert Options** button should be displayed. Read/write **Boolean** .|
|[DisplayNoteIndicator](application-displaynoteindicator-property-excel.md)| **True** if cells containing notes display cell tips and contain note indicators (small dots in their upper-right corners). Read/write **Boolean** .|
|[DisplayPasteOptions](application-displaypasteoptions-property-excel.md)| **True** if the **Paste Options** button can be displayed. Read/write **Boolean** .|
|[DisplayRecentFiles](application-displayrecentfiles-property-excel.md)| **True** if the list of recently used files is displayed in the UI. Read/write **Boolean** .|
|[DisplayScrollBars](application-displayscrollbars-property-excel.md)| **True** if scroll bars are visible for all workbooks. Read/write **Boolean** .|
|[DisplayStatusBar](application-displaystatusbar-property-excel.md)| **True** if the status bar is displayed. Read/write **Boolean** .|
|[EditDirectlyInCell](application-editdirectlyincell-property-excel.md)| **True** if Microsoft Excel allows editing in cells. Read/write **Boolean** .|
|[EnableAutoComplete](application-enableautocomplete-property-excel.md)| **True** if the AutoComplete feature is enabled. Read/write **Boolean** .|
|[EnableCancelKey](application-enablecancelkey-property-excel.md)|Controls how Microsoft Excel handles CTRL+BREAK (or ESC or COMMAND+PERIOD) user interruptions to the running procedure. Read/write  **[XlEnableCancelKey](xlenablecancelkey-enumeration-excel.md)** .|
|[EnableCheckFileExtensions](application-enablecheckfileextensions-property-excel.md)||
|[EnableEvents](application-enableevents-property-excel.md)| **True** if events are enabled for the specified object. Read/write **Boolean** .|
|[EnableLargeOperationAlert](application-enablelargeoperationalert-property-excel.md)|Sets or returns a  **Boolean** that represents whether to display an alert message when a user attempts to perform an operation that affects a larger number of cells than is specified in the Office center UI. Read/write **Boolean** .|
|[EnableLivePreview](application-enablelivepreview-property-excel.md)|Sets or returns a  **Boolean** that represents whether to show or hide gallery previews that appear when using galleries that support previewing. Setting this property to **True** shows a preview of your workbook before applying the command. Read/write **Boolean** .|
|[EnableMacroAnimations](application-enablemacroanimations-property-excel.md)|Controls whether macro animations are enabled.  **True** if user interface animations or chart animations are enabled. Is set to **False** (no animation) by default. If it is set to **True** during the running of a macro, it will enable animation and then will reset to **False** after the macro runs. Read/write **Boolean** .|
|[EnableSound](application-enablesound-property-excel.md)| **True** if sound is enabled for Microsoft Office. Read/write **Boolean** .|
|[ErrorCheckingOptions](application-errorcheckingoptions-property-excel.md)|Returns an  **[ErrorCheckingOptions](errorcheckingoptions-object-excel.md)** object, which represents the error checking options for an application.|
|[Excel4IntlMacroSheets](application-excel4intlmacrosheets-property-excel.md)|Returns a  **[Sheets](sheets-object-excel.md)** collection that represents all the Microsoft Excel 4.0 international macro sheets in the specified workbook. Read-only.|
|[Excel4MacroSheets](application-excel4macrosheets-property-excel.md)|Returns a  **[Sheets](sheets-object-excel.md)** collection that represents all the Microsoft Excel 4.0 macro sheets in the specified workbook. Read-only.|
|[ExtendList](application-extendlist-property-excel.md)| **True** if Microsoft Excel automatically extends formatting and formulas to new data that is added to a list. Read/write **Boolean** .|
|[FeatureInstall](application-featureinstall-property-excel.md)|Returns or sets a value (constant) that specifies how Microsoft Excel handles calls to methods and properties that require features that aren't yet installed. Can be one of the  **[MsoFeatureInstall](msofeatureinstall-enumeration-office.md)** constants listed in the following table. Read/write **MsoFeatureInstall** .|
|[FileConverters](application-fileconverters-property-excel.md)|Returns information about installed file converters. Returns  **null** if there are no converters installed. Read-only **Variant** .|
|[FileDialog](application-filedialog-property-excel.md)|Returns a  **[FileDialog](filedialog-object-office.md)** object representing an instance of the file dialog.|
|[FileExportConverters](application-fileexportconverters-property-excel.md)|Returns a  **[FileExportConverters](fileexportconverters-object-excel.md)** collection that represents all the file converters for saving files available to Microsoft Excel. Read-only.|
|[FileValidation](application-filevalidation-property-excel.md)|Returns or sets how Excel will validate files before opening them. Read/write|
|[FileValidationPivot](application-filevalidationpivot-property-excel.md)|Returns or sets how Excel will validate the contents of the data caches for PivotTable reports. Read/write|
|[FindFormat](application-findformat-property-excel.md)|Sets or returns the search criteria for the type of cell formats to find.|
|[FixedDecimal](application-fixeddecimal-property-excel.md)|All data entered after this property is set to  **True** will be formatted with the number of fixed decimal places set by the **[FixedDecimalPlaces](application-fixeddecimalplaces-property-excel.md)** property. Read/write **Boolean** .|
|[FixedDecimalPlaces](application-fixeddecimalplaces-property-excel.md)|Returns or sets the number of fixed decimal places used when the  **[FixedDecimal](application-fixeddecimal-property-excel.md)** property is set to **True** . Read/write **Long** .|
|[FlashFill](application-flashfill-property-excel.md)| **TRUE** indicates that the Excel Flash Fill feature has been enabled and active. **Boolean** Read/Write|
|[FlashFillMode](application-flashfillmode-property-excel.md)| **True** if the Flash Fill feature is enabled. **Boolean** Read/Write|
|[FormulaBarHeight](application-formulabarheight-property-excel.md)|Allows the user to specify the height of the formula bar in lines. Read/write  **Long** .|
|[GenerateGetPivotData](application-generategetpivotdata-property-excel.md)|Returns  **True** when Microsoft Excel can get PivotTable report data. Read/write **Boolean** .|
|[GenerateTableRefs](application-generatetablerefs-property-excel.md)|The  **GenerateTableRefs** property determines whether the traditional notation method or the new structured referencing notation method is used for referencing tables in formulas. Read/write.|
|[Height](application-height-property-excel.md)|Returns or sets a  **Double** value that represents tThe height, in points, of the main application window.|
|[HighQualityModeForGraphics](application-highqualitymodeforgraphics-property-excel.md)|Returns or sets whether Excel uses high quality mode to print graphics. Read/write|
|[Hinstance](application-hinstance-property-excel.md)|Returns a handle to the instance of Microsoft Excel 2013 represented by the [Application](application-object-excel.md) object. Read-only **Long** .|
|[HinstancePtr](application-hinstanceptr-property-excel.md)|Returns a handle to the instance of Microsoft Excel 2013 represented by the specified  **[Application](application-object-excel.md)** object. Read-only **Variant** .|
|[Hwnd](application-hwnd-property-excel.md)|Returns a  **Long** indicating the top-level window handle of the Microsoft Excel window. Read-only.|
|[IgnoreRemoteRequests](application-ignoreremoterequests-property-excel.md)| **True** if remote DDE requests are ignored. Read/write **Boolean** .|
|[Interactive](application-interactive-property-excel.md)| **True** if Microsoft Excel is in interactive mode; this property is usually **True** . If you set the this property to **False** , Microsoft Excel will block all input from the keyboard and mouse (except input to dialog boxes that are displayed by your code). Read/write **Boolean** .|
|[International](application-international-property-excel.md)|Returns information about the current country/region and international settings. Read-only  **Variant** .|
|[IsSandboxed](application-issandboxed-property-excel.md)|Returns  **True** if the specified workbook is open in a **Protected View** window. Read-only|
|[Iteration](application-iteration-property-excel.md)| **True** if Microsoft Excel will use iteration to resolve circular references. Read/write **Boolean** .|
|[LanguageSettings](application-languagesettings-property-excel.md)|Returns the  **[LanguageSettings](languagesettings-object-office.md)** object, which contains information about the language settings in Microsoft Excel. Read-only.|
|[LargeOperationCellThousandCount](application-largeoperationcellthousandcount-property-excel.md)|Returns or sets the maximum number of cells needed in an operation beyond which an alert is triggered. Read/write  **Long** .|
|[Left](application-left-property-excel.md)|Returns or sets a  **Double** value that represents the distance, in points, from the left edge of the screen to the left edge of the main Microsoft Excel window.|
|[LibraryPath](application-librarypath-property-excel.md)|Returns the path to the Library folder, but without the final separator. Read-only  **String** .|
|[MailSession](application-mailsession-property-excel.md)|Returns the MAPI mail session number as a hexadecimal string (if there's an active session), or returns  **null** if there's no session. Read-only **Variant** .|
|[MailSystem](application-mailsystem-property-excel.md)|Returns the mail system that's installed on the host machine. Read-only  **[XlMailSystem](xlmailsystem-enumeration-excel.md)** .|
|[MapPaperSize](application-mappapersize-property-excel.md)| **True** if documents formatted for the standard paper size of another country/region (for example, A4) are automatically adjusted so that they're printed correctly on the standard paper size (for example, Letter) of your country/region. Read/write **Boolean** .|
|[MathCoprocessorAvailable](application-mathcoprocessoravailable-property-excel.md)| **True** if a math coprocessor is available. Read-only **Boolean** .|
|[MaxChange](application-maxchange-property-excel.md)|Returns or sets the maximum amount of change between each iteration as Microsoft Excel resolves circular references. Read/write  **Double** .|
|[MaxIterations](application-maxiterations-property-excel.md)|Returns or sets the maximum number of iterations that Microsoft Excel can use to resolve a circular reference. Read/write  **Long** .|
|[MeasurementUnit](application-measurementunit-property-excel.md)|Specifies the measurement unit used in the application. Read/write  **xlMeasurementUnit** .|
|[MergeInstances](application-mergeinstances-property-excel.md)||
|[MouseAvailable](application-mouseavailable-property-excel.md)| **True** if a mouse is available. Read-only **Boolean** .|
|[MoveAfterReturn](application-moveafterreturn-property-excel.md)| **True** if the active cell will be moved as soon as the ENTER (RETURN) key is pressed. Read/write **Boolean** .|
|[MoveAfterReturnDirection](application-moveafterreturndirection-property-excel.md)|Returns or sets the direction in which the active cell is moved when the user presses ENTER. Read/write  **[XlDirection](xldirection-enumeration-excel.md)** .|
|[MultiThreadedCalculation](application-multithreadedcalculation-property-excel.md)|Returns a  **MultiThreadedCalculation** object that controls the multi-threaded recalculation settings. Read-only.|
|[Name](application-name-property-excel.md)|Returns a  **String** value that represents the name of the object.|
|[Names](application-names-property-excel.md)|Returns a  **[Names](names-object-excel.md)** collection that represents all the names in the active workbook. Read-only **Names** object.|
|[NetworkTemplatesPath](application-networktemplatespath-property-excel.md)|Returns the network path where templates are stored. If the network path doesn't exist, this property returns an empty string. Read-only  **String** .|
|[NewWorkbook](application-newworkbook-property-excel.md)|Returns a  **[NewFile](newfile-object-office.md)** object.|
|[ODBCErrors](application-odbcerrors-property-excel.md)|Returns an  **[ODBCErrors](application-odbcerrors-property-excel.md)** collection that contains all the ODBC errors generated by the most recent query table or PivotTable report operation. Read-only.|
|[ODBCTimeout](application-odbctimeout-property-excel.md)|Returns or sets the ODBC query time limit, in seconds. The default value is 45 seconds. Read/write  **Long** .|
|[OLEDBErrors](application-oledberrors-property-excel.md)|Returns the  **[OLEDBErrors](oledberrors-object-excel.md)** collection, which represents the error information returned by the most recent OLE DB query. Read-only.|
|[OnWindow](application-onwindow-property-excel.md)|Returns or sets the name of the procedure that's run whenever you activate a window. Read/write  **String** .|
|[OperatingSystem](application-operatingsystem-property-excel.md)|Returns the name and version number of the current operating system ??? for example, "Windows (32-bit) 4.00" or "Macintosh 7.00". Read-only  **String** .|
|[OrganizationName](application-organizationname-property-excel.md)|Returns the registered organization name. Read-only  **String** .|
|[Parent](application-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[Path](application-path-property-excel.md)|Returns a  **String** value that represents the complete path to the application, excluding the final separator and name of the application.|
|[PathSeparator](application-pathseparator-property-excel.md)|Returns the path separator character ("\"). Read-only  **String** .|
|[PivotTableSelection](application-pivottableselection-property-excel.md)| **True** if PivotTable reports use structured selection. Read/write **Boolean** .|
|[PreviousSelections](application-previousselections-property-excel.md)|Returns an array of the last four ranges or names selected. Each element in the array is a  **[Range](range-object-excel.md)** object. Read-only **Variant** .|
|[PrintCommunication](application-printcommunication-property-excel.md)|Specifies whether communication with the printer is turned on.  **Boolean** Read/write|
|[ProductCode](application-productcode-property-excel.md)|Returns the globally unique identifier (GUID) for Microsoft Excel. Read-only  **String** .|
|[PromptForSummaryInfo](application-promptforsummaryinfo-property-excel.md)| **True** if Microsoft Excel asks for summary information when files are first saved. Read/write **Boolean** .|
|[ProtectedViewWindows](application-protectedviewwindows-property-excel.md)|Returns a  **[ProtectedViewWindows](protectedviewwindows-object-excel.md)** collection that represents all the **Protected View** windows that are open in the application. Read-only|
|[QuickAnalysis](application-quickanalysis-property-excel.md)|Returns a  **[QuickAnalysis](quickanalysis-object-excel.md)** object that represents the Quick Analysis options of the application.|
|[Range](application-range-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents a cell or a range of cells.|
|[Ready](application-ready-property-excel.md)|Returns  **True** when the Microsoft Excel application is ready; **False** when the Excel application is not ready. Read-only **Boolean** .|
|[RecentFiles](application-recentfiles-property-excel.md)|Returns a  **[RecentFiles](recentfiles-object-excel.md)** collection that represents the list of recently used files.|
|[RecordRelative](application-recordrelative-property-excel.md)| **True** if macros are recorded using relative references; **False** if recording is absolute. Read-only **Boolean** .|
|[ReferenceStyle](application-referencestyle-property-excel.md)|Returns or sets how Microsoft Excel displays cell references and row and column headings in either A1 or R1C1 reference style. Read/write  **[XlReferenceStyle](xlreferencestyle-enumeration-excel.md)** .|
|[RegisteredFunctions](application-registeredfunctions-property-excel.md)|Returns information about functions in either dynamic-link libraries (DLLs) or code resources that were registered with the REGISTER or REGISTER.ID macro functions. Read-only  **Variant** .|
|[ReplaceFormat](application-replaceformat-property-excel.md)|Sets the replacement criteria to use in replacing cell formats. The replacement criteria is then used in a subsequent call to the Replace method of the Range object.|
|[RollZoom](application-rollzoom-property-excel.md)| **True** if the IntelliMouse zooms instead of scrolling. Read/write **Boolean** .|
|[Rows](application-rows-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents all the rows on the active worksheet. If the active document isn't a worksheet, the **Rows** property fails. Read-only **Range** object.|
|[RTD](application-rtd-property-excel.md)|Returns an  **[RTD](rtd-object-excel.md)** object.|
|[ScreenUpdating](application-screenupdating-property-excel.md)| **True** if screen updating is turned on. Read/write **Boolean** .|
|[Selection](application-selection-property-excel.md)|Returns the selected object in the active window for an  **[Application](application-object-excel.md)** object.|
|[Sheets](application-sheets-property-excel.md)|Returns a  **[Sheets](sheets-object-excel.md)** collection that represents all the sheets in the active workbook. Read-only **Sheets** object.|
|[SheetsInNewWorkbook](application-sheetsinnewworkbook-property-excel.md)|Returns or sets the number of sheets that Microsoft Excel automatically inserts into new workbooks. Read/write  **Long** .|
|[ShowChartTipNames](application-showcharttipnames-property-excel.md)| **True** if charts show chart tip names. The default value is **True** . Read/write **Boolean** .|
|[ShowChartTipValues](application-showcharttipvalues-property-excel.md)| **True** if charts show chart tip values. The default value is **True** . Read/write **Boolean** .|
|[ShowDevTools](application-showdevtools-property-excel.md)|Returns or sets a  **Boolean** that represents whether the **Developer** tab is displayed in the ribbon. Read/write **Boolean** .|
|[ShowMenuFloaties](application-showmenufloaties-property-excel.md)|Returns or sets a  **Boolean** that represents whether to display **Mini toolbars** when the user right-clicks in the workbook window. **False** if **Mini toolbars** are displayed. Read/write **Boolean** .|
|[ShowQuickAnalysis](application-showquickanalysis-property-excel.md)|Controls whether the Quick Analysis contextual user interface is displayed on selection.  **TRUE** means the Quick Analysis button will show. Corresponds to the **Show Quick Analysis options on selection** checkbox located in the **File** menu, **Options**,  **Excel Options**, and then  **General** tab. Read/Write. **Boolean** .|
|[ShowSelectionFloaties](application-showselectionfloaties-property-excel.md)|Returns or sets a  **Boolean** that represents whether **Mini toolbars** displays when a user selects text. **False** if **Mini toolbars** are displayed. Read/write **Boolean** .|
|[ShowStartupDialog](application-showstartupdialog-property-excel.md)|Returns  **True** (default is **False** ) when the New Workbook task pane appears for a Microsoft Excel application. Read/write **Boolean** .|
|[ShowToolTips](application-showtooltips-property-excel.md)| **True** if ToolTips are turned on. Read/write **Boolean** .|
|[SmartArtColors](application-smartartcolors-property-excel.md)|Returns the set of color styles that are currently loaded in the application. Read-only|
|[SmartArtLayouts](application-smartartlayouts-property-excel.md)|Returns the set of SmartArt layouts that are currently loaded in the application. Read-only|
|[SmartArtQuickStyles](application-smartartquickstyles-property-excel.md)|Returns the set of SmartArt quick styles which are currently loaded in the application. Read-only|
|[Speech](application-speech-property-excel.md)|Returns a  **[Speech](speech-object-excel.md)** object.|
|[SpellingOptions](application-spellingoptions-property-excel.md)|Returns a  **[SpellingOptions](spellingoptions-object-excel.md)** object that represents the spelling options of the application.|
|[StandardFont](application-standardfont-property-excel.md)|Returns or sets the name of the standard font. Read/write  **String** .|
|[StandardFontSize](application-standardfontsize-property-excel.md)|Returns or sets the standard font size, in points. Read/write  **Long** .|
|[StartupPath](application-startuppath-property-excel.md)|Returns the complete path of the startup folder, excluding the final separator. Read-only  **String** .|
|[StatusBar](application-statusbar-property-excel.md)|Returns or sets the text in the status bar. Read/write  **String** .|
|[TemplatesPath](application-templatespath-property-excel.md)|Returns the local path where templates are stored. Read-only  **String** .|
|[ThisCell](application-thiscell-property-excel.md)|Returns the cell in which the user-defined function is being called from as a  **[Range](range-object-excel.md)** object.|
|[ThisWorkbook](application-thisworkbook-property-excel.md)|Returns a  **[Workbook](workbook-object-excel.md)** object that represents the workbook where the current macro code is running. Read-only.|
|[ThousandsSeparator](application-thousandsseparator-property-excel.md)|Sets or returns the character used for the thousands separator as a  **String** . Read/write.|
|[Top](application-top-property-excel.md)|Returns or sets a  **Double** value that represents the distance, in points, from the top edge of the screen to the top edge of the main Microsoft Excel window.|
|[TransitionMenuKey](application-transitionmenukey-property-excel.md)|Returns or sets the Microsoft Excel menu or help key, which is usually "/". Read/write  **String** .|
|[TransitionMenuKeyAction](application-transitionmenukeyaction-property-excel.md)|Returns or sets the action taken when the Microsoft Excel menu key is pressed. Can be either  **xlExcelMenus** or **xlLotusHelp** . Read/write **Long** .|
|[TransitionNavigKeys](application-transitionnavigkeys-property-excel.md)| **True** if transition navigation keys are active. Read/write **Boolean** .|
|[UsableHeight](application-usableheight-property-excel.md)|Returns the maximum height of the space that a window can occupy in the application window area, in points. Read-only  **Double** .|
|[UsableWidth](application-usablewidth-property-excel.md)|Returns the maximum width of the space that a window can occupy in the application window area, in points. Read-only  **Double** .|
|[UseClusterConnector](application-useclusterconnector-property-excel.md)|Returns or sets whether Excel allows user-defined functions in XLL add-ins to be run on a compute cluster. Read/write|
|[UsedObjects](application-usedobjects-property-excel.md)|Returns a [UsedObjects](usedobjects-object-excel.md)object representing objects allocated in a workbook. Read-only|
|[UserControl](application-usercontrol-property-excel.md)| **True** if the application is visible or if it was created or started by the user. **False** if you created or started the application programmatically by using the **CreateObject** or **GetObject** functions, and the application is hidden. Read/write **Boolean** .|
|[UserLibraryPath](application-userlibrarypath-property-excel.md)|Returns the path to the location on the user's computer where the COM add-ins are installed. Read-only  **String** .|
|[UserName](application-username-property-excel.md)|Returns or sets the name of the current user. Read/write  **String** .|
|[UseSystemSeparators](application-usesystemseparators-property-excel.md)| **True** (default) if the system separators of Microsoft Excel are enabled. Read/write **Boolean** .|
|[Value](application-value-property-excel.md)|Returns a  **String** value that represents the name of the application.|
|[VBE](application-vbe-property-excel.md)|Returns a  **VBE** object that represents the Visual Basic Editor. Read-only.|
|[Version](application-version-property-excel.md)|Returns a  **String** value that represents the Microsoft Excel version number.|
|[Visible](application-visible-property-excel.md)|Returns or sets a  **Boolean** value that determines whether the object is visible. Read/write.|
|[WarnOnFunctionNameConflict](application-warnonfunctionnameconflict-property-excel.md)|The  **WarnOnFunctionNameConflict** property, when set to **True** , raises an alert if a developer tries to create a new function using an existing function name. Read/write **Boolean** .|
|[Watches](application-watches-property-excel.md)|Returns a  **[Watches](watches-object-excel.md)** object representing a range which is tracked when the worksheet is recalculated.|
|[Width](application-width-property-excel.md)|Returns or sets a  **Double** value that represents the distance, in points, from the left edge of the application window to its right edge.|
|[Windows](application-windows-property-excel.md)|Returns a  **[Windows](windows-object-excel.md)** collection that represents all the windows in all the workbooks. Read-only **Windows** object.|
|[WindowsForPens](application-windowsforpens-property-excel.md)| **True** if the computer is running under Microsoft Windows for Pen Computing. Read-only **Boolean** .|
|[WindowState](application-windowstate-property-excel.md)|Returns or sets the state of the window. Read/write  **[XlWindowState](xlwindowstate-enumeration-excel.md)** .|
|[Workbooks](application-workbooks-property-excel.md)|Returns a  **[Workbooks](workbooks-object-excel.md)** collection that represents all the open workbooks. Read-only.|
|[WorksheetFunction](application-worksheetfunction-property-excel.md)|Returns the  **[WorksheetFunction](worksheetfunction-object-excel.md)** object. Read-only.|
|[Worksheets](application-worksheets-property-excel.md)|For an  **Application** object, returns a **[Sheets](sheets-object-excel.md)** collection that represents all the worksheets in the active workbook. For a **Workbook** object, returns a **[Sheets](sheets-object-excel.md)** collection that represents all the worksheets in the specified workbook. Read-only **Sheets** object.|
|[EnableAnimations](application-enableanimations-property-excel.md)|This object, member, or enumeration is deprecated and is not intended to be used in your code. |

