---
title: Range Object (Excel)
keywords: vbaxl10.chm143072
f1_keywords:
- vbaxl10.chm143072
ms.prod: EXCEL
api_name:
- Excel.Range
ms.assetid: b8207778-0dcc-4570-1234-f130532cc8cd
---


# Range Object (Excel)

Represents a cell, a row, a column, a selection of cells containing one or more contiguous blocks of cells, or a 3-D range.


## Example

Use  **Range** ( _arg_ ), where _arg_ names the range, to return a **Range** object that represents a single cell or a range of cells. The following example places the value of cell A1 in cell A5.


```
Worksheets("Sheet1").Range("A5").Value = _ 
    Worksheets("Sheet1").Range("A1").Value
```

The following example fills the range A1:H8 with random numbers by setting the formula for each cell in the range. When it's used without an object qualifier (an object to the left of the period), the  **Range** property returns a range on the active sheet. If the active sheet isn't a worksheet, the method fails. Use the **[Activate](http://msdn.microsoft.com/library/worksheet-activate-method-excel%28Office.15%29.aspx)** method to activate a worksheet before you use the **Range** property without an explicit object qualifier.




```
Worksheets("Sheet1").Activate 
Range("A1:H8").Formula = "=Rand()"    'Range is on the active sheet
```

The following example clears the contents of the range named  _Criteria_.


 **Note**  If you use a text argument for the range address, you must specify the address in A1-style notation (you cannot use R1C1-style notation).




```
Worksheets(1).Range("Criteria").ClearContents
```

Use  **Cells** ( _row_, _column_ ) where _row_ is the row index and _column_ is the column index, to return a single cell. The following example sets the value of cell A1 to 24.




```
Worksheets(1).Cells(1, 1).Value = 24
```

The following example sets the formula for cell A2.




```
ActiveSheet.Cells(2, 1).Formula = "=Sum(B1:B5)"
```

Although you can also use  `Range("A1")` to return cell A1, there may be times when the **Cells** property is more convenient because you can use a variable for the row or column. The following example creates column and row headings on Sheet1. Be aware that after the worksheet has been activated, the **Cells** property can be used without an explicit sheet declaration (it returns a cell on the active sheet).


 **Note**  Although you could use Visual Basic string functions to alter A1-style references, it is easier (and better programming practice) to use the  `Cells(1, 1)` notation.




```
Sub SetUpTable() 
Worksheets("Sheet1").Activate 
For TheYear = 1 To 5 
    Cells(1, TheYear + 1).Value = 1990 + TheYear 
Next TheYear 
For TheQuarter = 1 To 4 
    Cells(TheQuarter + 1, 1).Value = "Q" &amp; TheQuarter 
Next TheQuarter 
End Sub
```

Use  _expression_. **Cells** ( _row_, _column_ ), where _expression_ is an expression that returns a **Range** object, and _row_ and _column_ are relative to the upper-left corner of the range, to return part of a range. The following example sets the formula for cell C5.




```
Worksheets(1).Range("C5:C10").Cells(1, 1).Formula = "=Rand()"
```

Use  **Range** ( _cell1, cell2_ ), where _cell1_ and _cell2_ are **Range** objects that specify the start and end cells, to return a **Range** object. The following example sets the border line style for cells A1:J10.


 **Note**  Be aware that the period in front of each occurrence of the  **Cells** property. The period is required if the result of the preceding **With** statement is to be applied to the **Cells** property—in this case, to indicate that the cells are on worksheet one (without the period, the **Cells** property would return cells on the active sheet).




```
With Worksheets(1) 
    .Range(.Cells(1, 1), _ 
        .Cells(10, 10)).Borders.LineStyle = xlThick 
End With
```

Use  **Offset** ( _row, column_ ), where _row_ and _column_ are the row and column offsets, to return a range at a specified offset to another range. The following example selects the cell three rows down from and one column to the right of the cell in the upper-left corner of the current selection. You cannot select a cell that is not on the active sheet, so you must first activate the worksheet.




```
Worksheets("Sheet1").Activate 
  'Can't select unless the sheet is active 
Selection.Offset(3, 1).Range("A1").Select
```

Use  **Union** ( _range1, range2_, ...) to return multiple-area ranges—that is, ranges composed of two or more contiguous blocks of cells. The following example creates an object defined as the union of ranges A1:B2 and C3:D4, and then selects the defined range.




```
Dim r1 As Range, r2 As Range, myMultiAreaRange As Range 
Worksheets("sheet1").Activate 
Set r1 = Range("A1:B2") 
Set r2 = Range("C3:D4") 
Set myMultiAreaRange = Union(r1, r2) 
myMultiAreaRange.Select
```

If you work with selections that contain more than one area, the  **[Areas](http://msdn.microsoft.com/library/range-areas-property-excel%28Office.15%29.aspx)** property is useful. It divides a multiple-area selection into individual **Range** objects and then returns the objects as a collection. You can use the **[Count](http://msdn.microsoft.com/library/range-count-property-excel%28Office.15%29.aspx)** property on the returned collection to verify a selection that contains more than one area, as shown in the following example.




```
Sub NoMultiAreaSelection() 
    NumberOfSelectedAreas = Selection.Areas.Count 
    If NumberOfSelectedAreas > 1 Then 
        MsgBox "You cannot carry out this command " &amp; _ 
            "on multi-area selections" 
    End If 
End Sub
```

 **Sample code provided by:** Dennis Wallentin,[VSTO &amp; .NET &amp; Excel](http://xldennis.wordpress.com/)

This example uses the  **AdvancedFilter** method of the **Range** object to create a list of the unique values, and the number of times those unique values occur, in the range of column A.




```
Sub Create_Unique_List_Count()
    'Excel workbook, the source and target worksheets, and the source and target ranges.
    Dim wbBook As Workbook
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim rnSource As Range
    Dim rnTarget As Range
    Dim rnUnique As Range
    'Variant to hold the unique data
    Dim vaUnique As Variant
    'Number of unique values in the data
    Dim lnCount As Long
    
    'Initialize the Excel objects
    Set wbBook = ThisWorkbook
    With wbBook
        Set wsSource = .Worksheets("Sheet1")
        Set wsTarget = .Worksheets("Sheet2")
    End With
    
    'On the source worksheet, set the range to the data stored in column A
    With wsSource
        Set rnSource = .Range(.Range("A1"), .Range("A100").End(xlDown))
    End With
    
    'On the target worksheet, set the range as column A.
    Set rnTarget = wsTarget.Range("A1")
    
    'Use AdvancedFilter to copy the data from the source to the target,
    'while filtering for duplicate values.
    rnSource.AdvancedFilter Action:=xlFilterCopy, _
                            CopyToRange:=rnTarget, _
                            Unique:=True
                            
    'On the target worksheet, set the unique range on Column A, excluding the first cell
    '(which will contain the "List" header for the column).
    With wsTarget
        Set rnUnique = .Range(.Range("A2"), .Range("A100").End(xlUp))
    End With
    
    'Assign all the values of the Unique range into the Unique variant.
    vaUnique = rnUnique.Value
    
    'Count the number of occurrences of every unique value in the source data,
    'and list it next to its relevant value.
    For lnCount = 1 To UBound(vaUnique)
        rnUnique(lnCount, 1).Offset(0, 1).Value = _
            Application.Evaluate("COUNTIF(" &amp; _
            rnSource.Address(External:=True) &amp; _
            ",""" &amp; rnUnique(lnCount, 1).Text &amp; """)")
    Next lnCount
    
    'Label the column of occurrences with "Occurrences"
    With rnTarget.Offset(0, 1)
        .Value = "Occurrences"
        .Font.Bold = True
    End With

End Sub
```


## Remarks

The following properties and methods for returning a  **Range** object are described in the examples section:


-  **[Range](http://msdn.microsoft.com/library/worksheet-range-property-excel%28Office.15%29.aspx)** property
    
-  **[Cells](http://msdn.microsoft.com/library/worksheet-cells-property-excel%28Office.15%29.aspx)** property
    
-  **Range** and **Cells**
    
-  **[Offset](http://msdn.microsoft.com/library/range-offset-property-excel%28Office.15%29.aspx)** property
    
-  **[Union](http://msdn.microsoft.com/library/application-union-method-excel%28Office.15%29.aspx)** method
    

## Methods



|**Name**|
|:-----|
|[Activate](http://msdn.microsoft.com/library/range-activate-method-excel%28Office.15%29.aspx)|
|[AddComment](http://msdn.microsoft.com/library/range-addcomment-method-excel%28Office.15%29.aspx)|
|[AdvancedFilter](http://msdn.microsoft.com/library/range-advancedfilter-method-excel%28Office.15%29.aspx)|
|[AllocateChanges](http://msdn.microsoft.com/library/range-allocatechanges-method-excel%28Office.15%29.aspx)|
|[ApplyNames](http://msdn.microsoft.com/library/range-applynames-method-excel%28Office.15%29.aspx)|
|[ApplyOutlineStyles](http://msdn.microsoft.com/library/range-applyoutlinestyles-method-excel%28Office.15%29.aspx)|
|[AutoComplete](http://msdn.microsoft.com/library/range-autocomplete-method-excel%28Office.15%29.aspx)|
|[AutoFill](http://msdn.microsoft.com/library/range-autofill-method-excel%28Office.15%29.aspx)|
|[AutoFilter](http://msdn.microsoft.com/library/range-autofilter-method-excel%28Office.15%29.aspx)|
|[AutoFit](http://msdn.microsoft.com/library/range-autofit-method-excel%28Office.15%29.aspx)|
|[AutoOutline](http://msdn.microsoft.com/library/range-autooutline-method-excel%28Office.15%29.aspx)|
|[BorderAround](http://msdn.microsoft.com/library/range-borderaround-method-excel%28Office.15%29.aspx)|
|[Calculate](http://msdn.microsoft.com/library/range-calculate-method-excel%28Office.15%29.aspx)|
|[CalculateRowMajorOrder](http://msdn.microsoft.com/library/range-calculaterowmajororder-method-excel%28Office.15%29.aspx)|
|[CheckSpelling](http://msdn.microsoft.com/library/range-checkspelling-method-excel%28Office.15%29.aspx)|
|[Clear](http://msdn.microsoft.com/library/range-clear-method-excel%28Office.15%29.aspx)|
|[ClearComments](http://msdn.microsoft.com/library/range-clearcomments-method-excel%28Office.15%29.aspx)|
|[ClearContents](http://msdn.microsoft.com/library/range-clearcontents-method-excel%28Office.15%29.aspx)|
|[ClearFormats](http://msdn.microsoft.com/library/range-clearformats-method-excel%28Office.15%29.aspx)|
|[ClearHyperlinks](http://msdn.microsoft.com/library/range-clearhyperlinks-method-excel%28Office.15%29.aspx)|
|[ClearNotes](http://msdn.microsoft.com/library/range-clearnotes-method-excel%28Office.15%29.aspx)|
|[ClearOutline](http://msdn.microsoft.com/library/range-clearoutline-method-excel%28Office.15%29.aspx)|
|[ColumnDifferences](http://msdn.microsoft.com/library/range-columndifferences-method-excel%28Office.15%29.aspx)|
|[Consolidate](http://msdn.microsoft.com/library/range-consolidate-method-excel%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/range-copy-method-excel%28Office.15%29.aspx)|
|[CopyFromRecordset](http://msdn.microsoft.com/library/range-copyfromrecordset-method-excel%28Office.15%29.aspx)|
|[CopyPicture](http://msdn.microsoft.com/library/range-copypicture-method-excel%28Office.15%29.aspx)|
|[CreateNames](http://msdn.microsoft.com/library/range-createnames-method-excel%28Office.15%29.aspx)|
|[Cut](http://msdn.microsoft.com/library/range-cut-method-excel%28Office.15%29.aspx)|
|[DataSeries](http://msdn.microsoft.com/library/range-dataseries-method-excel%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/range-delete-method-excel%28Office.15%29.aspx)|
|[DialogBox](http://msdn.microsoft.com/library/range-dialogbox-method-excel%28Office.15%29.aspx)|
|[Dirty](http://msdn.microsoft.com/library/range-dirty-method-excel%28Office.15%29.aspx)|
|[DiscardChanges](http://msdn.microsoft.com/library/range-discardchanges-method-excel%28Office.15%29.aspx)|
|[EditionOptions](http://msdn.microsoft.com/library/range-editionoptions-method-excel%28Office.15%29.aspx)|
|[ExportAsFixedFormat](http://msdn.microsoft.com/library/range-exportasfixedformat-method-excel%28Office.15%29.aspx)|
|[FillDown](http://msdn.microsoft.com/library/range-filldown-method-excel%28Office.15%29.aspx)|
|[FillLeft](http://msdn.microsoft.com/library/range-fillleft-method-excel%28Office.15%29.aspx)|
|[FillRight](http://msdn.microsoft.com/library/range-fillright-method-excel%28Office.15%29.aspx)|
|[FillUp](http://msdn.microsoft.com/library/range-fillup-method-excel%28Office.15%29.aspx)|
|[Find](http://msdn.microsoft.com/library/range-find-method-excel%28Office.15%29.aspx)|
|[FindNext](http://msdn.microsoft.com/library/range-findnext-method-excel%28Office.15%29.aspx)|
|[FindPrevious](http://msdn.microsoft.com/library/range-findprevious-method-excel%28Office.15%29.aspx)|
|[FlashFill](http://msdn.microsoft.com/library/range-flashfill-method-excel%28Office.15%29.aspx)|
|[FunctionWizard](http://msdn.microsoft.com/library/range-functionwizard-method-excel%28Office.15%29.aspx)|
|[Group](http://msdn.microsoft.com/library/range-group-method-excel%28Office.15%29.aspx)|
|[Insert](http://msdn.microsoft.com/library/range-insert-method-excel%28Office.15%29.aspx)|
|[InsertIndent](http://msdn.microsoft.com/library/range-insertindent-method-excel%28Office.15%29.aspx)|
|[Justify](http://msdn.microsoft.com/library/range-justify-method-excel%28Office.15%29.aspx)|
|[ListNames](http://msdn.microsoft.com/library/range-listnames-method-excel%28Office.15%29.aspx)|
|[Merge](http://msdn.microsoft.com/library/range-merge-method-excel%28Office.15%29.aspx)|
|[NavigateArrow](http://msdn.microsoft.com/library/range-navigatearrow-method-excel%28Office.15%29.aspx)|
|[NoteText](http://msdn.microsoft.com/library/range-notetext-method-excel%28Office.15%29.aspx)|
|[Parse](http://msdn.microsoft.com/library/range-parse-method-excel%28Office.15%29.aspx)|
|[PasteSpecial](http://msdn.microsoft.com/library/range-pastespecial-method-excel%28Office.15%29.aspx)|
|[PrintOut](http://msdn.microsoft.com/library/range-printout-method-excel%28Office.15%29.aspx)|
|[PrintPreview](http://msdn.microsoft.com/library/range-printpreview-method-excel%28Office.15%29.aspx)|
|[RemoveDuplicates](http://msdn.microsoft.com/library/range-removeduplicates-method-excel%28Office.15%29.aspx)|
|[RemoveSubtotal](http://msdn.microsoft.com/library/range-removesubtotal-method-excel%28Office.15%29.aspx)|
|[Replace](http://msdn.microsoft.com/library/range-replace-method-excel%28Office.15%29.aspx)|
|[RowDifferences](http://msdn.microsoft.com/library/range-rowdifferences-method-excel%28Office.15%29.aspx)|
|[Run](http://msdn.microsoft.com/library/range-run-method-excel%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/range-select-method-excel%28Office.15%29.aspx)|
|[SetPhonetic](http://msdn.microsoft.com/library/range-setphonetic-method-excel%28Office.15%29.aspx)|
|[Show](http://msdn.microsoft.com/library/range-show-method-excel%28Office.15%29.aspx)|
|[ShowDependents](http://msdn.microsoft.com/library/range-showdependents-method-excel%28Office.15%29.aspx)|
|[ShowErrors](http://msdn.microsoft.com/library/range-showerrors-method-excel%28Office.15%29.aspx)|
|[ShowPrecedents](http://msdn.microsoft.com/library/range-showprecedents-method-excel%28Office.15%29.aspx)|
|[Sort](http://msdn.microsoft.com/library/range-sort-method-excel%28Office.15%29.aspx)|
|[SortSpecial](http://msdn.microsoft.com/library/range-sortspecial-method-excel%28Office.15%29.aspx)|
|[Speak](http://msdn.microsoft.com/library/range-speak-method-excel%28Office.15%29.aspx)|
|[SpecialCells](http://msdn.microsoft.com/library/range-specialcells-method-excel%28Office.15%29.aspx)|
|[SubscribeTo](http://msdn.microsoft.com/library/range-subscribeto-method-excel%28Office.15%29.aspx)|
|[Subtotal](http://msdn.microsoft.com/library/range-subtotal-method-excel%28Office.15%29.aspx)|
|[Table](http://msdn.microsoft.com/library/range-table-method-excel%28Office.15%29.aspx)|
|[TextToColumns](http://msdn.microsoft.com/library/range-texttocolumns-method-excel%28Office.15%29.aspx)|
|[Ungroup](http://msdn.microsoft.com/library/range-ungroup-method-excel%28Office.15%29.aspx)|
|[UnMerge](http://msdn.microsoft.com/library/range-unmerge-method-excel%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AddIndent](http://msdn.microsoft.com/library/range-addindent-property-excel%28Office.15%29.aspx)|
|[Address](http://msdn.microsoft.com/library/range-address-property-excel%28Office.15%29.aspx)|
|[AddressLocal](http://msdn.microsoft.com/library/range-addresslocal-property-excel%28Office.15%29.aspx)|
|[AllowEdit](http://msdn.microsoft.com/library/range-allowedit-property-excel%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/range-application-property-excel%28Office.15%29.aspx)|
|[Areas](http://msdn.microsoft.com/library/range-areas-property-excel%28Office.15%29.aspx)|
|[Borders](http://msdn.microsoft.com/library/range-borders-property-excel%28Office.15%29.aspx)|
|[Cells](http://msdn.microsoft.com/library/range-cells-property-excel%28Office.15%29.aspx)|
|[Characters](http://msdn.microsoft.com/library/range-characters-property-excel%28Office.15%29.aspx)|
|[Column](http://msdn.microsoft.com/library/range-column-property-excel%28Office.15%29.aspx)|
|[Columns](http://msdn.microsoft.com/library/range-columns-property-excel%28Office.15%29.aspx)|
|[ColumnWidth](http://msdn.microsoft.com/library/range-columnwidth-property-excel%28Office.15%29.aspx)|
|[Comment](http://msdn.microsoft.com/library/range-comment-property-excel%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/range-count-property-excel%28Office.15%29.aspx)|
|[CountLarge](http://msdn.microsoft.com/library/range-countlarge-property-excel%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/range-creator-property-excel%28Office.15%29.aspx)|
|[CurrentArray](http://msdn.microsoft.com/library/range-currentarray-property-excel%28Office.15%29.aspx)|
|[CurrentRegion](http://msdn.microsoft.com/library/range-currentregion-property-excel%28Office.15%29.aspx)|
|[Dependents](http://msdn.microsoft.com/library/range-dependents-property-excel%28Office.15%29.aspx)|
|[DirectDependents](http://msdn.microsoft.com/library/range-directdependents-property-excel%28Office.15%29.aspx)|
|[DirectPrecedents](http://msdn.microsoft.com/library/range-directprecedents-property-excel%28Office.15%29.aspx)|
|[DisplayFormat](http://msdn.microsoft.com/library/range-displayformat-property-excel%28Office.15%29.aspx)|
|[End](http://msdn.microsoft.com/library/range-end-property-excel%28Office.15%29.aspx)|
|[EntireColumn](http://msdn.microsoft.com/library/range-entirecolumn-property-excel%28Office.15%29.aspx)|
|[EntireRow](http://msdn.microsoft.com/library/range-entirerow-property-excel%28Office.15%29.aspx)|
|[Errors](http://msdn.microsoft.com/library/range-errors-property-excel%28Office.15%29.aspx)|
|[Font](http://msdn.microsoft.com/library/range-font-property-excel%28Office.15%29.aspx)|
|[FormatConditions](http://msdn.microsoft.com/library/range-formatconditions-property-excel%28Office.15%29.aspx)|
|[Formula](http://msdn.microsoft.com/library/range-formula-property-excel%28Office.15%29.aspx)|
|[FormulaArray](http://msdn.microsoft.com/library/range-formulaarray-property-excel%28Office.15%29.aspx)|
|[FormulaHidden](http://msdn.microsoft.com/library/range-formulahidden-property-excel%28Office.15%29.aspx)|
|[FormulaLocal](http://msdn.microsoft.com/library/range-formulalocal-property-excel%28Office.15%29.aspx)|
|[FormulaR1C1](http://msdn.microsoft.com/library/range-formular1c1-property-excel%28Office.15%29.aspx)|
|[FormulaR1C1Local](http://msdn.microsoft.com/library/range-formular1c1local-property-excel%28Office.15%29.aspx)|
|[HasArray](http://msdn.microsoft.com/library/range-hasarray-property-excel%28Office.15%29.aspx)|
|[HasFormula](http://msdn.microsoft.com/library/range-hasformula-property-excel%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/range-height-property-excel%28Office.15%29.aspx)|
|[Hidden](http://msdn.microsoft.com/library/range-hidden-property-excel%28Office.15%29.aspx)|
|[HorizontalAlignment](http://msdn.microsoft.com/library/range-horizontalalignment-property-excel%28Office.15%29.aspx)|
|[Hyperlinks](http://msdn.microsoft.com/library/range-hyperlinks-property-excel%28Office.15%29.aspx)|
|[ID](http://msdn.microsoft.com/library/range-id-property-excel%28Office.15%29.aspx)|
|[IndentLevel](http://msdn.microsoft.com/library/range-indentlevel-property-excel%28Office.15%29.aspx)|
|[Interior](http://msdn.microsoft.com/library/range-interior-property-excel%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/range-item-property-excel%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/range-left-property-excel%28Office.15%29.aspx)|
|[ListHeaderRows](http://msdn.microsoft.com/library/range-listheaderrows-property-excel%28Office.15%29.aspx)|
|[ListObject](http://msdn.microsoft.com/library/range-listobject-property-excel%28Office.15%29.aspx)|
|[LocationInTable](http://msdn.microsoft.com/library/range-locationintable-property-excel%28Office.15%29.aspx)|
|[Locked](http://msdn.microsoft.com/library/range-locked-property-excel%28Office.15%29.aspx)|
|[MDX](http://msdn.microsoft.com/library/range-mdx-property-excel%28Office.15%29.aspx)|
|[MergeArea](http://msdn.microsoft.com/library/range-mergearea-property-excel%28Office.15%29.aspx)|
|[MergeCells](http://msdn.microsoft.com/library/range-mergecells-property-excel%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/range-name-property-excel%28Office.15%29.aspx)|
|[Next](http://msdn.microsoft.com/library/range-next-property-excel%28Office.15%29.aspx)|
|[NumberFormat](http://msdn.microsoft.com/library/range-numberformat-property-excel%28Office.15%29.aspx)|
|[NumberFormatLocal](http://msdn.microsoft.com/library/range-numberformatlocal-property-excel%28Office.15%29.aspx)|
|[Offset](http://msdn.microsoft.com/library/range-offset-property-excel%28Office.15%29.aspx)|
|[Orientation](http://msdn.microsoft.com/library/range-orientation-property-excel%28Office.15%29.aspx)|
|[OutlineLevel](http://msdn.microsoft.com/library/range-outlinelevel-property-excel%28Office.15%29.aspx)|
|[PageBreak](http://msdn.microsoft.com/library/range-pagebreak-property-excel%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/range-parent-property-excel%28Office.15%29.aspx)|
|[Phonetic](http://msdn.microsoft.com/library/range-phonetic-property-excel%28Office.15%29.aspx)|
|[Phonetics](http://msdn.microsoft.com/library/range-phonetics-property-excel%28Office.15%29.aspx)|
|[PivotCell](http://msdn.microsoft.com/library/range-pivotcell-property-excel%28Office.15%29.aspx)|
|[PivotField](http://msdn.microsoft.com/library/range-pivotfield-property-excel%28Office.15%29.aspx)|
|[PivotItem](http://msdn.microsoft.com/library/range-pivotitem-property-excel%28Office.15%29.aspx)|
|[PivotTable](http://msdn.microsoft.com/library/range-pivottable-property-excel%28Office.15%29.aspx)|
|[Precedents](http://msdn.microsoft.com/library/range-precedents-property-excel%28Office.15%29.aspx)|
|[PrefixCharacter](http://msdn.microsoft.com/library/range-prefixcharacter-property-excel%28Office.15%29.aspx)|
|[Previous](http://msdn.microsoft.com/library/range-previous-property-excel%28Office.15%29.aspx)|
|[QueryTable](http://msdn.microsoft.com/library/range-querytable-property-excel%28Office.15%29.aspx)|
|[Range](http://msdn.microsoft.com/library/range-range-property-excel%28Office.15%29.aspx)|
|[ReadingOrder](http://msdn.microsoft.com/library/range-readingorder-property-excel%28Office.15%29.aspx)|
|[Resize](http://msdn.microsoft.com/library/range-resize-property-excel%28Office.15%29.aspx)|
|[Row](http://msdn.microsoft.com/library/range-row-property-excel%28Office.15%29.aspx)|
|[RowHeight](http://msdn.microsoft.com/library/range-rowheight-property-excel%28Office.15%29.aspx)|
|[Rows](http://msdn.microsoft.com/library/range-rows-property-excel%28Office.15%29.aspx)|
|[ServerActions](http://msdn.microsoft.com/library/range-serveractions-property-excel%28Office.15%29.aspx)|
|[ShowDetail](http://msdn.microsoft.com/library/range-showdetail-property-excel%28Office.15%29.aspx)|
|[ShrinkToFit](http://msdn.microsoft.com/library/range-shrinktofit-property-excel%28Office.15%29.aspx)|
|[SoundNote](http://msdn.microsoft.com/library/range-soundnote-property-excel%28Office.15%29.aspx)|
|[SparklineGroups](http://msdn.microsoft.com/library/range-sparklinegroups-property-excel%28Office.15%29.aspx)|
|[Style](http://msdn.microsoft.com/library/range-style-property-excel%28Office.15%29.aspx)|
|[Summary](http://msdn.microsoft.com/library/range-summary-property-excel%28Office.15%29.aspx)|
|[Text](http://msdn.microsoft.com/library/range-text-property-excel%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/range-top-property-excel%28Office.15%29.aspx)|
|[UseStandardHeight](http://msdn.microsoft.com/library/range-usestandardheight-property-excel%28Office.15%29.aspx)|
|[UseStandardWidth](http://msdn.microsoft.com/library/range-usestandardwidth-property-excel%28Office.15%29.aspx)|
|[Validation](http://msdn.microsoft.com/library/range-validation-property-excel%28Office.15%29.aspx)|
|[Value](http://msdn.microsoft.com/library/range-value-property-excel%28Office.15%29.aspx)|
|[Value2](http://msdn.microsoft.com/library/range-value2-property-excel%28Office.15%29.aspx)|
|[VerticalAlignment](http://msdn.microsoft.com/library/range-verticalalignment-property-excel%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/range-width-property-excel%28Office.15%29.aspx)|
|[Worksheet](http://msdn.microsoft.com/library/range-worksheet-property-excel%28Office.15%29.aspx)|
|[WrapText](http://msdn.microsoft.com/library/range-wraptext-property-excel%28Office.15%29.aspx)|
|[XPath](http://msdn.microsoft.com/library/range-xpath-property-excel%28Office.15%29.aspx)|

## About the Contributor
<a name="AboutContributor"> </a>

Dennis Wallentin is the author of VSTO &amp; .NET &amp; Excel, a blog that focuses on .NET Framework solutions for Excel and Excel Services. Dennis has been developing Excel solutions for over 20 years and is also the co-author of "Professional Excel Development: The Definitive Guide to Developing Applications Using Microsoft Excel, VBA and .NET (2nd Edition)." 


