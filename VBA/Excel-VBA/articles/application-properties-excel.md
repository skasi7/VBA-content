---
title: Application Properties (Excel)
ms.prod: EXCEL
ms.assetid: 1826b882-f1c2-429d-a708-af334b8a6a8a
---


# Application Properties (Excel)

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
|[EnableCheckFileExtensions](application-enablecheckfileextensions-property-excel.md)| **True** to enable the **Tell me if Microsoft Excel isn't the default program for viewing and editing spreadsheets** dialog box. Read/write **Boolean** .|
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
|[MergeInstances](application-mergeinstances-property-excel.md)| **True** to merge multiple instances of the application into a single instance. Read/Write **Boolean** .|
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
|[OperatingSystem](application-operatingsystem-property-excel.md)|Returns the name and version number of the current operating system — for example, "Windows (32-bit) 4.00" or "Macintosh 7.00". Read-only  **String** .|
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

