---
title: Worksheet Object (Excel)
keywords: vbaxl10.chm173072
f1_keywords:
- vbaxl10.chm173072
ms.prod: EXCEL
api_name:
- Excel.Worksheet
ms.assetid: 182b705e-854a-81cc-a4b0-59b942de55ae
---


# Worksheet Object (Excel)

Represents a worksheet.


## Remarks

The  **Worksheet** object is a member of the **[Worksheets](http://msdn.microsoft.com/library/worksheets-object-excel%28Office.15%29.aspx)** collection. The **Worksheets** collection contains all the **Worksheet** objects in a workbook.

The  **Worksheet** object is also a member of the [Sheets](http://msdn.microsoft.com/library/sheets-object-excel%28Office.15%29.aspx) collection. The **Sheets** collection contains all the sheets in the workbook (both chart sheets and worksheets).


## Example

Use  **[Worksheets](http://msdn.microsoft.com/library/workbook-worksheets-property-excel%28Office.15%29.aspx)** (_index_), where _index_ is the worksheet index number or name, to return a single **Worksheet** object. The following example hides worksheet one in the active workbook.


```
Worksheets(1).Visible = False
```

The worksheet index number denotes the position of the worksheet on the workbook's tab bar.  `Worksheets(1)` is the first (leftmost) worksheet in the workbook, and `Worksheets(Worksheets.Count)` is the last one. All worksheets are included in the index count, even if they're hidden.



The worksheet name is shown on the tab for the worksheet. Use the [Name](http://msdn.microsoft.com/library/worksheet-name-property-excel%28Office.15%29.aspx) property to set or return the worksheet name. The following example protects the scenarios on Sheet1.




```
 
Dim strPassword As String 
strPassword = InputBox ("Enter the password for the worksheet") 
Worksheets("Sheet1").Protect password:=strPassword, scenarios:=True
```

When a worksheet is the active sheet, you can use the  [ActiveSheet](http://msdn.microsoft.com/library/workbook-activesheet-property-excel%28Office.15%29.aspx) property to refer to it. The following example uses the [Activate](http://msdn.microsoft.com/library/worksheet-activate-method-excel%28Office.15%29.aspx) method to activate Sheet1, sets the page orientation to landscape mode, and then prints the worksheet.




```
Worksheets("Sheet1").Activate 
ActiveSheet.PageSetup.Orientation = xlLandscape 
ActiveSheet.PrintOut
```

 **Sample code provided by:** Holy Macro! Books, [Holy Macro! It's 2,500 Excel VBA Examples](http://www.mrexcel.com/store/index.php?l=product_detail&amp;p=1)

This example uses the BeforeDoubleClick event to open a specified set of files in Notepad. To use this example your worksheet must contain the following data:


- Cell A1 must contain the names of the files to open, each separated by a comma and a space.
    
- Cell D1 must contain the path to where the Notepad files are located.
    
- Cell D2 must contain the path to where the Notepad program is located.
    
- Cell D3 must contain the file extension, without the period, for the Notepad files (txt).
    
When you double-click cell A1, the files specified in cell A1 are opened in Notepad.




```
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
   'Define your variables.
   Dim sFile As String, sPath As String, sTxt As String, sExe As String, sSfx As String
   
   'If you did not double-click on A1, then exit the function.
   If Target.Address <> "$A$1" Then Exit Sub
   
   'If you did double-click on A1, then override the default double-click behaviour with this function.
   Cancel = True
   
   'Set the path to the files, the path to Notepad, the file extension of the files, and the names of the files,
   'based on the information in the worksheet.
   sPath = Range("D1").Value
   sExe = Range("D2").Value
   sSfx = Range("D3").Value
   sFile = Range("A1").Value
   
   'Remove the spaces between the file names.
   sFile = WorksheetFunction.Substitute(sFile, " ", "")
   
   'Go through each file in the list (separated by commas) and
   'create the path, call the executable, and move on to the next comma.
   Do While InStr(sFile, ",")
      sTxt = sPath &amp; "\" &amp; Left(sFile, InStr(sFile, ",") - 1) &amp; "." &amp; sSfx
      If Dir(sTxt) <> "" Then Shell sExe &amp; " " &amp; sTxt, vbNormalFocus
      sFile = Right(sFile, Len(sFile) - InStr(sFile, ","))
   Loop
   
   'Finish off the last file name in the list
   sTxt = sPath &amp; "\" &amp; sFile &amp; "." &amp; sSfx
   If Dir(sTxt) <> "" Then Shell sExe &amp; " " &amp; sTxt, vbNormalNoFocus
End Sub
```


## Events



|**Name**|
|:-----|
|[Activate](http://msdn.microsoft.com/library/worksheet-activate-event-excel%28Office.15%29.aspx)|
|[BeforeDelete](http://msdn.microsoft.com/library/worksheet-beforedelete-event-excel%28Office.15%29.aspx)|
|[BeforeDoubleClick](http://msdn.microsoft.com/library/worksheet-beforedoubleclick-event-excel%28Office.15%29.aspx)|
|[BeforeRightClick](http://msdn.microsoft.com/library/worksheet-beforerightclick-event-excel%28Office.15%29.aspx)|
|[Calculate](http://msdn.microsoft.com/library/worksheet-calculate-event-excel%28Office.15%29.aspx)|
|[Change](http://msdn.microsoft.com/library/worksheet-change-event-excel%28Office.15%29.aspx)|
|[Deactivate](http://msdn.microsoft.com/library/worksheet-deactivate-event-excel%28Office.15%29.aspx)|
|[FollowHyperlink](http://msdn.microsoft.com/library/worksheet-followhyperlink-event-excel%28Office.15%29.aspx)|
|[LensGalleryRenderComplete](http://msdn.microsoft.com/library/worksheet-lensgalleryrendercomplete-event-excel%28Office.15%29.aspx)|
|[PivotTableAfterValueChange](http://msdn.microsoft.com/library/worksheet-pivottableaftervaluechange-event-excel%28Office.15%29.aspx)|
|[PivotTableBeforeAllocateChanges](http://msdn.microsoft.com/library/worksheet-pivottablebeforeallocatechanges-event-excel%28Office.15%29.aspx)|
|[PivotTableBeforeCommitChanges](http://msdn.microsoft.com/library/worksheet-pivottablebeforecommitchanges-event-excel%28Office.15%29.aspx)|
|[PivotTableBeforeDiscardChanges](http://msdn.microsoft.com/library/worksheet-pivottablebeforediscardchanges-event-excel%28Office.15%29.aspx)|
|[PivotTableChangeSync](http://msdn.microsoft.com/library/worksheet-pivottablechangesync-event-excel%28Office.15%29.aspx)|
|[PivotTableUpdate](http://msdn.microsoft.com/library/worksheet-pivottableupdate-event-excel%28Office.15%29.aspx)|
|[SelectionChange](http://msdn.microsoft.com/library/worksheet-selectionchange-event-excel%28Office.15%29.aspx)|
|[TableUpdate](http://msdn.microsoft.com/library/worksheet-tableupdate-event-excel%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Activate](http://msdn.microsoft.com/library/worksheet-activate-method-excel%28Office.15%29.aspx)|
|[Calculate](http://msdn.microsoft.com/library/worksheet-calculate-method-excel%28Office.15%29.aspx)|
|[ChartObjects](http://msdn.microsoft.com/library/worksheet-chartobjects-method-excel%28Office.15%29.aspx)|
|[CheckSpelling](http://msdn.microsoft.com/library/worksheet-checkspelling-method-excel%28Office.15%29.aspx)|
|[CircleInvalid](http://msdn.microsoft.com/library/worksheet-circleinvalid-method-excel%28Office.15%29.aspx)|
|[ClearArrows](http://msdn.microsoft.com/library/worksheet-cleararrows-method-excel%28Office.15%29.aspx)|
|[ClearCircles](http://msdn.microsoft.com/library/worksheet-clearcircles-method-excel%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/worksheet-copy-method-excel%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/worksheet-delete-method-excel%28Office.15%29.aspx)|
|[Evaluate](http://msdn.microsoft.com/library/worksheet-evaluate-method-excel%28Office.15%29.aspx)|
|[ExportAsFixedFormat](http://msdn.microsoft.com/library/worksheet-exportasfixedformat-method-excel%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/worksheet-move-method-excel%28Office.15%29.aspx)|
|[OLEObjects](http://msdn.microsoft.com/library/worksheet-oleobjects-method-excel%28Office.15%29.aspx)|
|[Paste](http://msdn.microsoft.com/library/worksheet-paste-method-excel%28Office.15%29.aspx)|
|[PasteSpecial](http://msdn.microsoft.com/library/worksheet-pastespecial-method-excel%28Office.15%29.aspx)|
|[PivotTables](http://msdn.microsoft.com/library/worksheet-pivottables-method-excel%28Office.15%29.aspx)|
|[PivotTableWizard](http://msdn.microsoft.com/library/worksheet-pivottablewizard-method-excel%28Office.15%29.aspx)|
|[PrintOut](http://msdn.microsoft.com/library/worksheet-printout-method-excel%28Office.15%29.aspx)|
|[PrintPreview](http://msdn.microsoft.com/library/worksheet-printpreview-method-excel%28Office.15%29.aspx)|
|[Protect](http://msdn.microsoft.com/library/worksheet-protect-method-excel%28Office.15%29.aspx)|
|[ResetAllPageBreaks](http://msdn.microsoft.com/library/worksheet-resetallpagebreaks-method-excel%28Office.15%29.aspx)|
|[SaveAs](http://msdn.microsoft.com/library/worksheet-saveas-method-excel%28Office.15%29.aspx)|
|[Scenarios](http://msdn.microsoft.com/library/worksheet-scenarios-method-excel%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/worksheet-select-method-excel%28Office.15%29.aspx)|
|[SetBackgroundPicture](http://msdn.microsoft.com/library/worksheet-setbackgroundpicture-method-excel%28Office.15%29.aspx)|
|[ShowAllData](http://msdn.microsoft.com/library/worksheet-showalldata-method-excel%28Office.15%29.aspx)|
|[ShowDataForm](http://msdn.microsoft.com/library/worksheet-showdataform-method-excel%28Office.15%29.aspx)|
|[Unprotect](http://msdn.microsoft.com/library/worksheet-unprotect-method-excel%28Office.15%29.aspx)|
|[XmlDataQuery](http://msdn.microsoft.com/library/worksheet-xmldataquery-method-excel%28Office.15%29.aspx)|
|[XmlMapQuery](http://msdn.microsoft.com/library/worksheet-xmlmapquery-method-excel%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/worksheet-application-property-excel%28Office.15%29.aspx)|
|[AutoFilter](http://msdn.microsoft.com/library/worksheet-autofilter-property-excel%28Office.15%29.aspx)|
|[AutoFilterMode](http://msdn.microsoft.com/library/worksheet-autofiltermode-property-excel%28Office.15%29.aspx)|
|[Cells](http://msdn.microsoft.com/library/worksheet-cells-property-excel%28Office.15%29.aspx)|
|[CircularReference](http://msdn.microsoft.com/library/worksheet-circularreference-property-excel%28Office.15%29.aspx)|
|[CodeName](http://msdn.microsoft.com/library/worksheet-codename-property-excel%28Office.15%29.aspx)|
|[Columns](http://msdn.microsoft.com/library/worksheet-columns-property-excel%28Office.15%29.aspx)|
|[Comments](http://msdn.microsoft.com/library/worksheet-comments-property-excel%28Office.15%29.aspx)|
|[ConsolidationFunction](http://msdn.microsoft.com/library/worksheet-consolidationfunction-property-excel%28Office.15%29.aspx)|
|[ConsolidationOptions](http://msdn.microsoft.com/library/worksheet-consolidationoptions-property-excel%28Office.15%29.aspx)|
|[ConsolidationSources](http://msdn.microsoft.com/library/worksheet-consolidationsources-property-excel%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/worksheet-creator-property-excel%28Office.15%29.aspx)|
|[CustomProperties](http://msdn.microsoft.com/library/worksheet-customproperties-property-excel%28Office.15%29.aspx)|
|[DisplayPageBreaks](http://msdn.microsoft.com/library/worksheet-displaypagebreaks-property-excel%28Office.15%29.aspx)|
|[DisplayRightToLeft](http://msdn.microsoft.com/library/worksheet-displayrighttoleft-property-excel%28Office.15%29.aspx)|
|[EnableAutoFilter](http://msdn.microsoft.com/library/worksheet-enableautofilter-property-excel%28Office.15%29.aspx)|
|[EnableCalculation](http://msdn.microsoft.com/library/worksheet-enablecalculation-property-excel%28Office.15%29.aspx)|
|[EnableFormatConditionsCalculation](http://msdn.microsoft.com/library/worksheet-enableformatconditionscalculation-property-excel%28Office.15%29.aspx)|
|[EnableOutlining](http://msdn.microsoft.com/library/worksheet-enableoutlining-property-excel%28Office.15%29.aspx)|
|[EnablePivotTable](http://msdn.microsoft.com/library/worksheet-enablepivottable-property-excel%28Office.15%29.aspx)|
|[EnableSelection](http://msdn.microsoft.com/library/worksheet-enableselection-property-excel%28Office.15%29.aspx)|
|[FilterMode](http://msdn.microsoft.com/library/worksheet-filtermode-property-excel%28Office.15%29.aspx)|
|[HPageBreaks](http://msdn.microsoft.com/library/worksheet-hpagebreaks-property-excel%28Office.15%29.aspx)|
|[Hyperlinks](http://msdn.microsoft.com/library/worksheet-hyperlinks-property-excel%28Office.15%29.aspx)|
|[Index](http://msdn.microsoft.com/library/worksheet-index-property-excel%28Office.15%29.aspx)|
|[ListObjects](http://msdn.microsoft.com/library/worksheet-listobjects-property-excel%28Office.15%29.aspx)|
|[MailEnvelope](http://msdn.microsoft.com/library/worksheet-mailenvelope-property-excel%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/worksheet-name-property-excel%28Office.15%29.aspx)|
|[Names](http://msdn.microsoft.com/library/worksheet-names-property-excel%28Office.15%29.aspx)|
|[Next](http://msdn.microsoft.com/library/worksheet-next-property-excel%28Office.15%29.aspx)|
|[Outline](http://msdn.microsoft.com/library/worksheet-outline-property-excel%28Office.15%29.aspx)|
|[PageSetup](http://msdn.microsoft.com/library/worksheet-pagesetup-property-excel%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/worksheet-parent-property-excel%28Office.15%29.aspx)|
|[Previous](http://msdn.microsoft.com/library/worksheet-previous-property-excel%28Office.15%29.aspx)|
|[PrintedCommentPages](http://msdn.microsoft.com/library/worksheet-printedcommentpages-property-excel%28Office.15%29.aspx)|
|[ProtectContents](http://msdn.microsoft.com/library/worksheet-protectcontents-property-excel%28Office.15%29.aspx)|
|[ProtectDrawingObjects](http://msdn.microsoft.com/library/worksheet-protectdrawingobjects-property-excel%28Office.15%29.aspx)|
|[Protection](http://msdn.microsoft.com/library/worksheet-protection-property-excel%28Office.15%29.aspx)|
|[ProtectionMode](http://msdn.microsoft.com/library/worksheet-protectionmode-property-excel%28Office.15%29.aspx)|
|[ProtectScenarios](http://msdn.microsoft.com/library/worksheet-protectscenarios-property-excel%28Office.15%29.aspx)|
|[QueryTables](http://msdn.microsoft.com/library/worksheet-querytables-property-excel%28Office.15%29.aspx)|
|[Range](http://msdn.microsoft.com/library/worksheet-range-property-excel%28Office.15%29.aspx)|
|[Rows](http://msdn.microsoft.com/library/worksheet-rows-property-excel%28Office.15%29.aspx)|
|[ScrollArea](http://msdn.microsoft.com/library/worksheet-scrollarea-property-excel%28Office.15%29.aspx)|
|[Shapes](http://msdn.microsoft.com/library/worksheet-shapes-property-excel%28Office.15%29.aspx)|
|[Sort](http://msdn.microsoft.com/library/worksheet-sort-property-excel%28Office.15%29.aspx)|
|[StandardHeight](http://msdn.microsoft.com/library/worksheet-standardheight-property-excel%28Office.15%29.aspx)|
|[StandardWidth](http://msdn.microsoft.com/library/worksheet-standardwidth-property-excel%28Office.15%29.aspx)|
|[Tab](http://msdn.microsoft.com/library/worksheet-tab-property-excel%28Office.15%29.aspx)|
|[TransitionExpEval](http://msdn.microsoft.com/library/worksheet-transitionexpeval-property-excel%28Office.15%29.aspx)|
|[TransitionFormEntry](http://msdn.microsoft.com/library/worksheet-transitionformentry-property-excel%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/worksheet-type-property-excel%28Office.15%29.aspx)|
|[UsedRange](http://msdn.microsoft.com/library/worksheet-usedrange-property-excel%28Office.15%29.aspx)|
|[Visible](http://msdn.microsoft.com/library/worksheet-visible-property-excel%28Office.15%29.aspx)|
|[VPageBreaks](http://msdn.microsoft.com/library/worksheet-vpagebreaks-property-excel%28Office.15%29.aspx)|

## About the Contributor
<a name="AboutContributor"> </a>

Holy Macro! Books publishes entertaining books for people who use Microsoft Office. See the complete catalog at MrExcel.com. 


## See also
<a name="AboutContributor"> </a>


#### Other resources

<<<<<<< HEAD
=======


>>>>>>> d7667e83d23dbf8ebf5bf068ba6fed14c840c0f5
[Excel Object Model Reference](http://msdn.microsoft.com/library/object-model-excel-vba-reference%28Office.15%29.aspx)

