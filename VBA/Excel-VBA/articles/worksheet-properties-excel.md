---
title: Worksheet Properties (Excel)
ms.prod: EXCEL
ms.assetid: e7914ee3-c864-4ef3-9ab1-c271c5c2741f
---


# Worksheet Properties (Excel)

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

