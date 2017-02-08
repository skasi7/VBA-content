---
title: Worksheet Members (Excel)
ms.prod: EXCEL
ms.assetid: f8c1afea-1a1c-f5e4-37e3-52c434c8c157
---


# Worksheet Members (Excel)
Represents a worksheet.

Represents a worksheet.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[Activate](worksheet-activate-event-excel.md)|Occurs when a workbook, worksheet, chart sheet, or embedded chart is activated.|
|[BeforeDelete](worksheet-beforedelete-event-excel.md)||
|[BeforeDoubleClick](worksheet-beforedoubleclick-event-excel.md)|Occurs when a worksheet is double-clicked, before the default double-click action.|
|[BeforeRightClick](worksheet-beforerightclick-event-excel.md)|Occurs when a worksheet is right-clicked, before the default right-click action.|
|[Calculate](worksheet-calculate-event-excel.md)|Occurs after the worksheet is recalculated, for the  **Worksheet** object.|
|[Change](worksheet-change-event-excel.md)|Occurs when cells on the worksheet are changed by the user or by an external link.|
|[Deactivate](worksheet-deactivate-event-excel.md)|Occurs when the chart, worksheet, or workbook is deactivated.|
|[FollowHyperlink](worksheet-followhyperlink-event-excel.md)|Occurs when you click any hyperlink on a worksheet. For application- and workbook-level events, see the  **[SheetFollowHyperlink](application-sheetfollowhyperlink-event-excel.md)** event and **[SheetFollowHyperlink](workbook-sheetfollowhyperlink-event-excel.md)** event.|
|[LensGalleryRenderComplete](worksheet-lensgalleryrendercomplete-event-excel.md)|Occurs when a callout gallery's icons (dynamic &; static) have completed rendering.|
|[PivotTableAfterValueChange](worksheet-pivottableaftervaluechange-event-excel.md)|Occurs after a cell or range of cells inside a PivotTable are edited or recalculated (for cells that contain formulas).|
|[PivotTableBeforeAllocateChanges](worksheet-pivottablebeforeallocatechanges-event-excel.md)|Occurs before changes are applied to a PivotTable.|
|[PivotTableBeforeCommitChanges](worksheet-pivottablebeforecommitchanges-event-excel.md)|Occurs before changes are committed against the OLAP data source for a PivotTable.|
|[PivotTableBeforeDiscardChanges](worksheet-pivottablebeforediscardchanges-event-excel.md)|Occurs before changes to a PivotTable are discarded.|
|[PivotTableChangeSync](worksheet-pivottablechangesync-event-excel.md)|Occurs after changes to a PivotTable.|
|[PivotTableUpdate](worksheet-pivottableupdate-event-excel.md)|Occurs after a PivotTable report is updated on a worksheet.|
|[SelectionChange](worksheet-selectionchange-event-excel.md)|Occurs when the selection changes on a worksheet.|
|[TableUpdate](worksheet-tableupdate-event-excel.md)|Occurs after a Query table connected to the Data Model is updated on a worksheet.|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Activate](worksheet-activate-method-excel.md)|Makes the current sheet the active sheet. |
|[Calculate](worksheet-calculate-method-excel.md)|Calculates all open workbooks, a specific worksheet in a workbook, or a specified range of cells on a worksheet, as shown in the following table.|
|[ChartObjects](worksheet-chartobjects-method-excel.md)|Returns an object that represents either a single embedded chart (a  **[ChartObject](chartobject-object-excel.md)** object) or a collection of all the embedded charts (a **[ChartObjects](chartobjects-object-excel.md)** object) on the sheet.|
|[CheckSpelling](worksheet-checkspelling-method-excel.md)|Checks the spelling of an object.|
|[CircleInvalid](worksheet-circleinvalid-method-excel.md)|Circles invalid entries on the worksheet.|
|[ClearArrows](worksheet-cleararrows-method-excel.md)|Clears the tracer arrows from the worksheet. Tracer arrows are added by using the auditing feature.|
|[ClearCircles](worksheet-clearcircles-method-excel.md)|Clears circles from invalid entries on the worksheet.|
|[Copy](worksheet-copy-method-excel.md)|Copies the sheet to another location in the workbook.|
|[Delete](worksheet-delete-method-excel.md)|Deletes the object.|
|[Evaluate](worksheet-evaluate-method-excel.md)|Converts a Microsoft Excel name to an object or a value.|
|[ExportAsFixedFormat](worksheet-exportasfixedformat-method-excel.md)|Exports to a file of the specified format.|
|[Move](worksheet-move-method-excel.md)|Moves the sheet to another location in the workbook.|
|[OLEObjects](worksheet-oleobjects-method-excel.md)|Returns an object that represents either a single OLE object (an  **[OLEObject](oleobject-object-excel.md)** ) or a collection of all OLE objects (an **[OLEObjects](oleobjects-object-excel.md)** collection) on the chart or sheet. Read-only.|
|[Paste](worksheet-paste-method-excel.md)|Pastes the contents of the Clipboard onto the sheet.|
|[PasteSpecial](worksheet-pastespecial-method-excel.md)|Pastes the contents of the Clipboard onto the sheet, using a specified format. Use this method to paste data from other applications or to paste data in a specific format.|
|[PivotTables](worksheet-pivottables-method-excel.md)|Returns an object that represents either a single PivotTable report (a  **[PivotTable](pivottable-object-excel.md)** object) or a collection of all the PivotTable reports (a **[PivotTables](pivottables-object-excel.md)** object) on a worksheet. Read-only.|
|[PivotTableWizard](worksheet-pivottablewizard-method-excel.md)|Creates a new PivotTable report. This method doesn't display the PivotTable Wizard. This method isn't available for OLE DB data sources. Use the  **[Add](pivottables-add-method-excel.md)** method to add a PivotTable cache, and then create a PivotTable report based on the cache.|
|[PrintOut](worksheet-printout-method-excel.md)|Prints the object.|
|[PrintPreview](worksheet-printpreview-method-excel.md)|Shows a preview of the object as it would look when printed.|
|[Protect](worksheet-protect-method-excel.md)|Protects a worksheet so that it cannot be modified.|
|[ResetAllPageBreaks](worksheet-resetallpagebreaks-method-excel.md)|Resets all page breaks on the specified worksheet.|
|[SaveAs](worksheet-saveas-method-excel.md)|Saves changes to the chart or worksheet in a different file.|
|[Scenarios](worksheet-scenarios-method-excel.md)|Returns an object that represents either a single scenario (a  **[Scenario](scenario-object-excel.md)** object) or a collection of scenarios (a **[Scenarios](scenarios-object-excel.md)** object) on the worksheet.|
|[Select](worksheet-select-method-excel.md)|Selects the object.|
|[SetBackgroundPicture](worksheet-setbackgroundpicture-method-excel.md)|Sets the background graphic for a worksheet.|
|[ShowAllData](worksheet-showalldata-method-excel.md)|Makes all rows of the currently filtered list visible. If AutoFilter is in use, this method changes the arrows to "All."|
|[ShowDataForm](worksheet-showdataform-method-excel.md)|Displays the data form associated with the worksheet.|
|[Unprotect](worksheet-unprotect-method-excel.md)|Removes protection from a sheet or workbook. This method has no effect if the sheet or workbook isn't protected.|
|[XmlDataQuery](worksheet-xmldataquery-method-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents the cells mapped to a particular XPath. Returns **Nothing** if the specified XPath has not been mapped to the worksheet, or if the mapped range is empty.|
|[XmlMapQuery](worksheet-xmlmapquery-method-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents the cells mapped to a particular XPath. Returns **Nothing** if the specified XPath has not been mapped to the worksheet.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](worksheet-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[AutoFilter](worksheet-autofilter-property-excel.md)|Returns an  **AutoFilter** object if filtering is on. Read-only.|
|[AutoFilterMode](worksheet-autofiltermode-property-excel.md)| **True** if the AutoFilter drop-down arrows are currently displayed on the sheet. This property is independent of the **FilterMode** property. Read/write **Boolean** .|
|[Cells](worksheet-cells-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents all the cells on the worksheet (not just the cells that are currently in use).|
|[CircularReference](worksheet-circularreference-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents the range containing the first circular reference on the sheet, or returns **Nothing** if there's no circular reference on the sheet. The circular reference must be removed before calculation can proceed.|
|[CodeName](worksheet-codename-property-excel.md)|Returns the code name for the object. Read-only  **String** .|
|[Columns](worksheet-columns-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents all the columns on the active worksheet. If the active document isn't a worksheet, the **Columns** property fails.|
|[Comments](worksheet-comments-property-excel.md)|Returns a  **[Comments](comments-object-excel.md)** collection that represents all the comments for the specified worksheet. Read-only.|
|[ConsolidationFunction](worksheet-consolidationfunction-property-excel.md)|Returns the function code used for the current consolidation. Can be one of the constants of  **[XlConsolidationFunction](xlconsolidationfunction-enumeration-excel.md)** . Read-only **Long** .|
|[ConsolidationOptions](worksheet-consolidationoptions-property-excel.md)|Returns a three-element array of consolidation options, as shown in the following table. If the element is  **True** , that option is set. Read-only **Variant** .|
|[ConsolidationSources](worksheet-consolidationsources-property-excel.md)|Returns an array of string values that name the source sheets for the worksheet's current consolidation. Returns  **Empty** if there's no consolidation on the sheet. Read-only **Variant** .|
|[Creator](worksheet-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[CustomProperties](worksheet-customproperties-property-excel.md)|Returns a  **[CustomProperties](customproperties-object-excel.md)** object representing the identifier information associated with a worksheet.|
|[DisplayPageBreaks](worksheet-displaypagebreaks-property-excel.md)| **True** if page breaks (both automatic and manual) on the specified worksheet are displayed. Read/write **Boolean** .|
|[DisplayRightToLeft](worksheet-displayrighttoleft-property-excel.md)| **True** if the specified worksheet is displayed from right to left instead of from left to right. **False** if the object is displayed from left to right. Read-only **Boolean** .|
|[EnableAutoFilter](worksheet-enableautofilter-property-excel.md)| **True** if AutoFilter arrows are enabled when user-interface-only protection is turned on. Read/write **Boolean** .|
|[EnableCalculation](worksheet-enablecalculation-property-excel.md)| **True** if Microsoft Excel automatically recalculates the worksheet when necessary. **False** if Excel doesn't recalculate the sheet. Read/write **Boolean** .|
|[EnableFormatConditionsCalculation](worksheet-enableformatconditionscalculation-property-excel.md)|Returms or sets if conditional formats will will occur automatically as needed. Read/write  **Boolean** .|
|[EnableOutlining](worksheet-enableoutlining-property-excel.md)| **True** if outlining symbols are enabled when user-interface-only protection is turned on. Read/write **Boolean** .|
|[EnablePivotTable](worksheet-enablepivottable-property-excel.md)| **True** if PivotTable controls and actions are enabled when user-interface-only protection is turned on. Read/write **Boolean** .|
|[EnableSelection](worksheet-enableselection-property-excel.md)|Returns or sets what can be selected on the sheet. Read/write  **[XlEnableSelection](xlenableselection-enumeration-excel.md)** .|
|[FilterMode](worksheet-filtermode-property-excel.md)| **True** if the worksheet is in the filter mode. Read-only **Boolean** .|
|[HPageBreaks](worksheet-hpagebreaks-property-excel.md)|Returns an  **[HPageBreaks](hpagebreaks-object-excel.md)** collection that represents the horizontal page breaks on the sheet. Read-only.|
|[Hyperlinks](worksheet-hyperlinks-property-excel.md)|Returns a  **[Hyperlinks](hyperlinks-object-excel.md)** collection that represents the hyperlinks for the worksheet.|
|[Index](worksheet-index-property-excel.md)|Returns a  **Long** value that represents the index number of the object within the collection of similar objects.|
|[ListObjects](worksheet-listobjects-property-excel.md)|Returns a collection of  **[ListObject](listobject-object-excel.md)** objects in the worksheet. Read-only **ListObjects** collection.|
|[MailEnvelope](worksheet-mailenvelope-property-excel.md)|Rrepresents an e-mail header for a document.|
|[Name](worksheet-name-property-excel.md)|Returns or sets a  **String** value that represents the object name.|
|[Names](worksheet-names-property-excel.md)|Returns a  **[Names](names-object-excel.md)** collection that represents all the worksheet-specific names (names defined with the "WorksheetName!" prefix). Read-only **Names** object.|
|[Next](worksheet-next-property-excel.md)|Returns a  **[Worksheet](worksheet-object-excel.md)** object that represents the next sheet.|
|[Outline](worksheet-outline-property-excel.md)|Returns an  **[Outline](outline-object-excel.md)** object that represents the outline for the specified worksheet. Read-only.|
|[PageSetup](worksheet-pagesetup-property-excel.md)|Returns a  **[PageSetup](pagesetup-object-excel.md)** object that contains all the page setup settings for the specified object. Read-only.|
|[Parent](worksheet-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[Previous](worksheet-previous-property-excel.md)|Returns a  **[Worksheet](worksheet-object-excel.md)** object that represents the next sheet.|
|[PrintedCommentPages](worksheet-printedcommentpages-property-excel.md)|Returns the number of comment pages that will be printed for the current worksheet. Read-only|
|[ProtectContents](worksheet-protectcontents-property-excel.md)| **True** if the contents of the sheet are protected. This protects the individual cells. To turn on content protection, use the **[Protect](worksheet-protect-method-excel.md)** method with the _Contents_ argument set to **True** . Read-only **Boolean** .|
|[ProtectDrawingObjects](worksheet-protectdrawingobjects-property-excel.md)| **True** if shapes are protected. To turn on shape protection, use the **[Protect](worksheet-protect-method-excel.md)** method with the _DrawingObjects_ argument set to **True** . Read-only **Boolean** .|
|[Protection](worksheet-protection-property-excel.md)|Returns a  **[Protection](protection-object-excel.md)** object that represents the protection options of the worksheet.|
|[ProtectionMode](worksheet-protectionmode-property-excel.md)| **True** if user-interface-only protection is turned on. To turn on user interface protection, use the **[Protect](worksheet-protect-method-excel.md)** method with the _UserInterfaceOnly_ argument set to **True** . Read-only **Boolean** .|
|[ProtectScenarios](worksheet-protectscenarios-property-excel.md)| **True** if the worksheet scenarios are protected. Read-only **Boolean** .|
|[QueryTables](worksheet-querytables-property-excel.md)|Returns the  **[QueryTables](querytables-object-excel.md)** collection that represents all the query tables on the specified worksheet. Read-only.|
|[Range](worksheet-range-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents a cell or a range of cells.|
|[Rows](worksheet-rows-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents all the rows on the specified worksheet. Read-only **Range** object.|
|[ScrollArea](worksheet-scrollarea-property-excel.md)|Returns or sets the range where scrolling is allowed, as an A1-style range reference. Cells outside the scroll area cannot be selected. Read/write  **String** .|
|[Shapes](worksheet-shapes-property-excel.md)|Returns a  **[Shapes](shapes-object-excel.md)** collection that represents all the shapes on the worksheet. Read-only.|
|[Sort](worksheet-sort-property-excel.md)|Returns a  **[Sort](sort-object-excel.md)** object. Read-only.|
|[StandardHeight](worksheet-standardheight-property-excel.md)|Returns the standard (default) height of all the rows in the worksheet, in points. Read-only  **Double** .|
|[StandardWidth](worksheet-standardwidth-property-excel.md)|Returns or sets the standard (default) width of all the columns in the worksheet. Read/write  **Double** .|
|[Tab](worksheet-tab-property-excel.md)|Returns a  **[Tab](tab-object-excel.md)** object for a worksheet.|
|[TransitionExpEval](worksheet-transitionexpeval-property-excel.md)| **True** if Microsoft Excel uses Lotus 1-2-3 expression evaluation rules for the worksheet. Read/write **Boolean** .|
|[TransitionFormEntry](worksheet-transitionformentry-property-excel.md)| **True** if Microsoft Excel uses Lotus 1-2-3 formula entry rules for the worksheet. Read/write **Boolean** .|
|[Type](worksheet-type-property-excel.md)|Returns an  **[XlSheetType](xlsheettype-enumeration-excel.md)** value that represents the worksheet type.|
|[UsedRange](worksheet-usedrange-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents the used range on the specified worksheet. Read-only.|
|[Visible](worksheet-visible-property-excel.md)|Returns or sets an  **[XlSheetVisibility](xlsheetvisibility-enumeration-excel.md)** value that determines whether the object is visible.|
|[VPageBreaks](worksheet-vpagebreaks-property-excel.md)|Returns a  **[VPageBreaks](worksheet-vpagebreaks-property-excel.md)** collection that represents the vertical page breaks on the sheet. Read-only.|

