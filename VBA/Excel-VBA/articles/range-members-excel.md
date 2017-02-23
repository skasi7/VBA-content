---
title: Range Members (Excel)
ms.prod: EXCEL
ms.assetid: 4336bf81-1e63-7e44-1792-baf366a027a7
---


# Range Members (Excel)
Represents a cell, a row, a column, a selection of cells containing one or more contiguous blocks of cells, or a 3-D range.

Represents a cell, a row, a column, a selection of cells containing one or more contiguous blocks of cells, or a 3-D range.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Activate](range-activate-method-excel.md)|Activates a single cell, which must be inside the current selection. To select a range of cells, use the  **[Select](range-select-method-excel.md)** method.|
|[AddComment](range-addcomment-method-excel.md)|Adds a comment to the range.|
|[AdvancedFilter](range-advancedfilter-method-excel.md)|Filters or copies data from a list based on a criteria range. If the initial selection is a single cell, that cell's current region is used.|
|[AllocateChanges](range-allocatechanges-method-excel.md)|Performs a writeback operation for all edited cells in a range based on an OLAP data source.|
|[ApplyNames](range-applynames-method-excel.md)|Applies names to the cells in the specified range.|
|[ApplyOutlineStyles](range-applyoutlinestyles-method-excel.md)|Applies outlining styles to the specified range.|
|[AutoComplete](range-autocomplete-method-excel.md)|Returns an AutoComplete match from the list. If there's no AutoComplete match or if more than one entry in the list matches the string to complete, this method returns an empty string.|
|[AutoFill](range-autofill-method-excel.md)|Performs an autofill on the cells in the specified range.|
|[AutoFilter](range-autofilter-method-excel.md)|Filters a list using the AutoFilter.|
|[AutoFit](range-autofit-method-excel.md)|Changes the width of the columns in the range or the height of the rows in the range to achieve the best fit.|
|[AutoOutline](range-autooutline-method-excel.md)|Automatically creates an outline for the specified range. If the range is a single cell, Microsoft Excel creates an outline for the entire sheet. The new outline replaces any existing outline.|
|[BorderAround](range-borderaround-method-excel.md)|Adds a border to a range and sets the  **[Color](border-color-property-excel.md)** , **[LineStyle](border-linestyle-property-excel.md)** , and **[Weight](border-weight-property-excel.md)** properties for the new border. **Variant** .|
|[Calculate](range-calculate-method-excel.md)|Calculates all open workbooks, a specific worksheet in a workbook, or a specified range of cells on a worksheet, as shown in the following table.|
|[CalculateRowMajorOrder](range-calculaterowmajororder-method-excel.md)|Calculates a specfied range of cells.|
|[CheckSpelling](range-checkspelling-method-excel.md)|Checks the spelling of an object.|
|[Clear](range-clear-method-excel.md)|Clears the entire object.|
|[ClearComments](range-clearcomments-method-excel.md)|Clears all cell comments from the specified range.|
|[ClearContents](range-clearcontents-method-excel.md)|Clears the formulas from the range.|
|[ClearFormats](range-clearformats-method-excel.md)|Clears the formatting of the object.|
|[ClearHyperlinks](range-clearhyperlinks-method-excel.md)|Removes all hyperlinks from the specified range.|
|[ClearNotes](range-clearnotes-method-excel.md)|Clears notes and sound notes from all the cells in the specified range.|
|[ClearOutline](range-clearoutline-method-excel.md)|Clears the outline for the specified range.|
|[ColumnDifferences](range-columndifferences-method-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents all the cells whose contents are different from the comparison cell in each column.|
|[Consolidate](range-consolidate-method-excel.md)|Consolidates data from multiple ranges on multiple worksheets into a single range on a single worksheet.  **Variant** .|
|[Copy](range-copy-method-excel.md)|Copies the range to the specified range or to the Clipboard.|
|[CopyFromRecordset](range-copyfromrecordset-method-excel.md)|Copies the contents of an ADO or DAO  **Recordset** object onto a worksheet, beginning at the upper-left corner of the specified range. If the **Recordset** object contains fields with OLE objects in them, this method fails.|
|[CopyPicture](range-copypicture-method-excel.md)|Copies the selected object to the Clipboard as a picture.  **Variant** .|
|[CreateNames](range-createnames-method-excel.md)|Creates names in the specified range, based on text labels in the sheet.|
|[Cut](range-cut-method-excel.md)|Cuts the object to the Clipboard or pastes it into a specified destination.|
|[DataSeries](range-dataseries-method-excel.md)|Creates a data series in the specified range.  **Variant** .|
|[Delete](range-delete-method-excel.md)|Deletes the object.|
|[DialogBox](range-dialogbox-method-excel.md)|Displays a dialog box defined by a dialog box definition table on a Microsoft Excel 4.0 macro sheet. Returns the number of the chosen control, or returns  **False** if the user clicks the **Cancel** button.|
|[Dirty](range-dirty-method-excel.md)|Designates a range to be recalculated when the next recalculation occurs.|
|[DiscardChanges](range-discardchanges-method-excel.md)|Discards all changes in the edited cells of the range.|
|[EditionOptions](range-editionoptions-method-excel.md)|You have requested Help for a Visual Basic keyword used only on the Macintosh. For information about this keyword, consult the language reference Help included with Microsoft Office Macintosh Edition.|
|[ExportAsFixedFormat](range-exportasfixedformat-method-excel.md)|Exports to a file of the specified format.|
|[FillDown](range-filldown-method-excel.md)|Fills down from the top cell or cells in the specified range to the bottom of the range. The contents and formatting of the cell or cells in the top row of a range are copied into the rest of the rows in the range.|
|[FillLeft](range-fillleft-method-excel.md)|Fills left from the rightmost cell or cells in the specified range. The contents and formatting of the cell or cells in the rightmost column of a range are copied into the rest of the columns in the range.|
|[FillRight](range-fillright-method-excel.md)|Fills right from the leftmost cell or cells in the specified range. The contents and formatting of the cell or cells in the leftmost column of a range are copied into the rest of the columns in the range.|
|[FillUp](range-fillup-method-excel.md)|Fills up from the bottom cell or cells in the specified range to the top of the range. The contents and formatting of the cell or cells in the bottom row of a range are copied into the rest of the rows in the range.|
|[Find](range-find-method-excel.md)|Finds specific information in a range.|
|[FindNext](range-findnext-method-excel.md)|Continues a search that was begun with the  **[Find](range-find-method-excel.md)** method. Finds the next cell that matches those same conditions and returns a **[Range](range-object-excel.md)** object that represents that cell. This does not affect the selection or the active cell.|
|[FindPrevious](range-findprevious-method-excel.md)|Continues a search that was begun with the  **[Find](range-find-method-excel.md)** method. Finds the previous cell that matches those same conditions and returns a **[Range](range-object-excel.md)** object that represents that cell. Doesn't affect the selection or the active cell.|
|[FlashFill](range-flashfill-method-excel.md)|TRUE indicates that the Excel Flash Fill feature has been enabled and active.|
|[FunctionWizard](range-functionwizard-method-excel.md)|Starts the Function Wizard for the upper-left cell of the range.|
|[Group](range-group-method-excel.md)|When the  **[Range](range-object-excel.md)** object represents a single cell in a PivotTable field's data range, the **Group** method performs numeric or date-based grouping in that field.|
|[Insert](range-insert-method-excel.md)|Inserts a cell or a range of cells into the worksheet or macro sheet and shifts other cells away to make space.|
|[InsertIndent](range-insertindent-method-excel.md)|Adds an indent to the specified range.|
|[Justify](range-justify-method-excel.md)|Rearranges the text in a range so that it fills the range evenly.|
|[ListNames](range-listnames-method-excel.md)|Pastes a list of all nonhidden names onto the worksheet, beginning with the first cell in the range.|
|[Merge](range-merge-method-excel.md)|Creates a merged cell from the specified  **[Range](range-object-excel.md)** object.|
|[NavigateArrow](range-navigatearrow-method-excel.md)|Navigates a tracer arrow for the specified range to the precedent, dependent, or error-causing cell or cells. Selects the precedent, dependent, or error cells and returns a  **[Range](range-object-excel.md)** object that represents the new selection. This method causes an error if it's applied to a cell without visible tracer arrows.|
|[NoteText](range-notetext-method-excel.md)|Returns or sets the cell note associated with the cell in the upper-left corner of the range. Read/write  **String** . Cell notes have been replaced by range comments. For more information, see the **[Comment](comment-object-excel.md)** object.|
|[Parse](range-parse-method-excel.md)|Parses a range of data and breaks it into multiple cells. Distributes the contents of the range to fill several adjacent columns; the range can be no more than one column wide.|
|[PasteSpecial](range-pastespecial-method-excel.md)|Pastes a  **[Range](range-object-excel.md)** from the Clipboard into the specified range.|
|[PrintOut](range-printout-method-excel.md)|Prints the object.|
|[PrintPreview](range-printpreview-method-excel.md)|Shows a preview of the object as it would look when printed.|
|[RemoveDuplicates](range-removeduplicates-method-excel.md)|Removes duplicate values from a range of values.|
|[RemoveSubtotal](range-removesubtotal-method-excel.md)|Removes subtotals from a list.|
|[Replace](range-replace-method-excel.md)|Returns a  **Boolean** indicating characters in cells within the specified range. Using this method doesn't change either the selection or the active cell.|
|[RowDifferences](range-rowdifferences-method-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents all the cells whose contents are different from those of the comparison cell in each row.|
|[Run](range-run-method-excel.md)|Runs the Microsoft Excel macro at this location. The range must be on a macro sheet.|
|[Select](range-select-method-excel.md)|Selects the object.|
|[SetPhonetic](range-setphonetic-method-excel.md)|Creates  **[Phonetic](phonetic-object-excel.md)** objects for all the cells in the specified range.|
|[Show](range-show-method-excel.md)|Scrolls through the contents of the active window to move the range into view. The range must consist of a single cell in the active document.|
|[ShowDependents](range-showdependents-method-excel.md)|Draws tracer arrows to the direct dependents of the range.|
|[ShowErrors](range-showerrors-method-excel.md)|Draws tracer arrows through the precedents tree to the cell that's the source of the error, and returns the range that contains that cell.|
|[ShowPrecedents](range-showprecedents-method-excel.md)|Draws tracer arrows to the direct precedents of the range.|
|[Sort](range-sort-method-excel.md)|Sorts a range of values.|
|[SortSpecial](range-sortspecial-method-excel.md)|Uses East Asian sorting methods to sort the range, a PivotTable report, or uses the method for the active region if the range contains only one cell. For example, Japanese sorts in the order of the Kana syllabary.|
|[Speak](range-speak-method-excel.md)|Causes the cells of the range to be spoken in row order or column order.|
|[SpecialCells](range-specialcells-method-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents all the cells that match the specified type and value.|
|[SubscribeTo](range-subscribeto-method-excel.md)|You have requested Help for a Visual Basic keyword used only on the Macintosh. For information about this keyword, consult the language reference Help included with Microsoft Office Macintosh Edition.|
|[Subtotal](range-subtotal-method-excel.md)|Creates subtotals for the range (or the current region, if the range is a single cell).|
|[Table](range-table-method-excel.md)|Creates a data table based on input values and formulas that you define on a worksheet.|
|[TextToColumns](range-texttocolumns-method-excel.md)|Parses a column of cells that contain text into several columns.|
|[Ungroup](range-ungroup-method-excel.md)|Promotes a range in an outline (that is, decreases its outline level). The specified range must be a row or column, or a range of rows or columns. If the range is in a PivotTable report, this method ungroups the items contained in the range.|
|[UnMerge](range-unmerge-method-excel.md)|Separates a merged area into individual cells.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[AddIndent](range-addindent-property-excel.md)|Returns or sets a  **Variant** value that indicates if text is automatically indented when the text alignment in a cell is set to equal distribution (either horizontally or vertically.)|
|[Address](range-address-property-excel.md)|Returns a  **String** value that represents the range reference in the language of the macro.|
|[AddressLocal](range-addresslocal-property-excel.md)|Returns the range reference for the specified range in the language of the user. Read-only  **String** .|
|[AllowEdit](range-allowedit-property-excel.md)|Returns a  **Boolean** value that indicates if the range can be edited on a protected worksheet.|
|[Application](range-application-property-excel.md)|When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.|
|[Areas](range-areas-property-excel.md)|Returns an  **[Areas](areas-object-excel.md)** collection that represents all the ranges in a multiple-area selection. Read-only.|
|[Borders](range-borders-property-excel.md)|Returns a  **[Borders](borders-object-excel.md)** collection that represents the borders of a style or a range of cells (including a range defined as part of a conditional format).|
|[Cells](range-cells-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents the cells in the specified range.|
|[Characters](range-characters-property-excel.md)|Returns a  **[Characters](characters-object-excel.md)** object that represents a range of characters within the object text. You can use the **Characters** object to format characters within a text string.|
|[Column](range-column-property-excel.md)|Returns the number of the first column in the first area in the specified range. Read-only  **Long** .|
|[Columns](range-columns-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents the columns in the specified range.|
|[ColumnWidth](range-columnwidth-property-excel.md)|Returns or sets the width of all columns in the specified range. Read/write  **Variant** .|
|[Comment](range-comment-property-excel.md)|Returns a  **[Comment](comment-object-excel.md)** object that represents the comment associated with the cell in the upper-left corner of the range.|
|[Count](range-count-property-excel.md)|Returns a  **Long** value that represents the number of objects in the collection.|
|[CountLarge](range-countlarge-property-excel.md)|Returns a value that represents the number of objects in the collection. Read-only  **Variant** .|
|[Creator](range-creator-property-excel.md)|Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .|
|[CurrentArray](range-currentarray-property-excel.md)|If the specified cell is part of an array, returns a  **[Range](range-object-excel.md)** object that represents the entire array. Read-only.|
|[CurrentRegion](range-currentregion-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents the current region. The current region is a range bounded by any combination of blank rows and blank columns. Read-only.|
|[Dependents](range-dependents-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents the range containing all the dependents of a cell. This can be a multiple selection (a union of **Range** objects) if there's more than one dependent. Read-only **Range** object.|
|[DirectDependents](range-directdependents-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents the range containing all the direct dependents of a cell. This can be a multiple selection (a union of **Range** objects) if there's more than one dependent. Read-only **Range** object.|
|[DirectPrecedents](range-directprecedents-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents the range containing all the direct precedents of a cell. This can be a multiple selection (a union of **Range** objects) if there's more than one precedent. Read-only **Range** object.|
|[DisplayFormat](range-displayformat-property-excel.md)|Returns a  **[DisplayFormat](displayformat-object-excel.md)** object that represents the display settings for the specified range. Read-only|
|[End](range-end-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents the cell at the end of the region that contains the source range. Equivalent to pressing END+UP ARROW, END+DOWN ARROW, END+LEFT ARROW, or END+RIGHT ARROW. Read-only **Range** object.|
|[EntireColumn](range-entirecolumn-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents the entire column (or columns) that contains the specified range. Read-only.|
|[EntireRow](range-entirerow-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents the entire row (or rows) that contains the specified range. Read-only.|
|[Errors](range-errors-property-excel.md)|Allows the user to to access error checking options.|
|[Font](range-font-property-excel.md)|Returns a  **[Font](font-object-excel.md)** object that represents the font of the specified object.|
|[FormatConditions](range-formatconditions-property-excel.md)|Returns a  **[FormatConditions](formatconditions-object-excel.md)** collection that represents all the conditional formats for the specified range. Read-only.|
|[Formula](range-formula-property-excel.md)|Returns or sets a  **Variant** value that represents the object's formula in A1-style notation and in the macro language.|
|[FormulaArray](range-formulaarray-property-excel.md)|Returns or sets the array formula of a range. Returns (or can be set to) a single formula or a Visual Basic array. If the specified range doesn't contain an array formula, this property returns  **null** . Read/write **Variant** .|
|[FormulaHidden](range-formulahidden-property-excel.md)|Returns or sets a  **Variant** value that indicates if the formula will be hidden when the worksheet is protected.|
|[FormulaLocal](range-formulalocal-property-excel.md)|Returns or sets the formula for the object, using A1-style references in the language of the user. Read/write  **Variant** .|
|[FormulaR1C1](range-formular1c1-property-excel.md)|Returns or sets the formula for the object, using R1C1-style notation in the language of the macro. Read/write  **Variant** .|
|[FormulaR1C1Local](range-formular1c1local-property-excel.md)|Returns or sets the formula for the object, using R1C1-style notation in the language of the user. Read/write  **Variant** .|
|[HasArray](range-hasarray-property-excel.md)| **True** if the specified cell is part of an array formula. Read-only **Variant** .|
|[HasFormula](range-hasformula-property-excel.md)| **True** if all cells in the range contain formulas; **False** if none of the cells in the range contains a formula; **null** otherwise. Read-only **Variant** .|
|[Height](range-height-property-excel.md)|Returns or sets a  **Variant** value that represents the height, in points, of the range.|
|[Hidden](range-hidden-property-excel.md)|Returns or sets a  **Variant** value that indicates if the rows or columns are hidden.|
|[HorizontalAlignment](range-horizontalalignment-property-excel.md)|Returns or sets a  **Variant** value that represents the horizontal alignment for the specified object.|
|[Hyperlinks](range-hyperlinks-property-excel.md)|Returns a  **[Hyperlinks](hyperlinks-object-excel.md)** collection that represents the hyperlinks for the range.|
|[ID](range-id-property-excel.md)|Returns or sets a  **String** value that represents the identifying label for the specified cell when the page is saved as a Web page.|
|[IndentLevel](range-indentlevel-property-excel.md)|Returns or sets a  **Variant** value that represents the indent level for the cell or range. Can be an integer from 0 to 15.|
|[Interior](range-interior-property-excel.md)|Returns an  **[Interior](interior-object-excel.md)** object that represents the interior of the specified object.|
|[Item](range-item-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents a range at an offset to the specified range.|
|[Left](range-left-property-excel.md)|Returns a  **Variant** value that represents the distance, in points, from the left edge of column A to the left edge of the range.|
|[ListHeaderRows](range-listheaderrows-property-excel.md)|Returns the number of header rows for the specified range. Read-only  **Long** .|
|[ListObject](range-listobject-property-excel.md)|Returns a  **[ListObject](listobject-object-excel.md)** object for the **[Range](range-object-excel.md)** object. Read-only **ListObject** object.|
|[LocationInTable](range-locationintable-property-excel.md)|Returns a constant that describes the part of the  **[PivotTable](pivottable-object-excel.md)** report that contains the upper-left corner of the specified range. Can be one of the following **[XlLocationInTable](xllocationintable-enumeration-excel.md)** . constants. Read-only **Long** .|
|[Locked](range-locked-property-excel.md)|Returns or sets a  **Variant** value that indicates if the object is locked.|
|[MDX](range-mdx-property-excel.md)|Returns the MDX name for the specified  **Range** object. Read-only **String** .|
|[MergeArea](range-mergearea-property-excel.md)|Returns a  **Range** object that represents the merged range containing the specified cell. If the specified cell isn't in a merged range, this property returns the specified cell. Read-only **Variant** .|
|[MergeCells](range-mergecells-property-excel.md)| **True** if the range contains merged cells. Read/write **Variant** .|
|[Name](range-name-property-excel.md)|Returns or sets a  **Variant** value that represents the name of the object.|
|[Next](range-next-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents the next cell.|
|[NumberFormat](range-numberformat-property-excel.md)|Returns or sets a  **Variant** value that represents the format code for the object.|
|[NumberFormatLocal](range-numberformatlocal-property-excel.md)|Returns or sets a  **Variant** value that represents the format code for the object as a string in the language of the user.|
|[Offset](range-offset-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents a range that's offset from the specified range.|
|[Orientation](range-orientation-property-excel.md)|Returns or sets a  **Variant** value that represents the text orientation.|
|[OutlineLevel](range-outlinelevel-property-excel.md)|Returns or sets the current outline level of the specified row or column. Read/write  **Variant** .|
|[PageBreak](range-pagebreak-property-excel.md)|Returns or sets the location of a page break. Can be one of the following  **[XlPageBreak](xlpagebreak-enumeration-excel.md)** constants: **xlPageBreakAutomatic** , **xlPageBreakManual** , or **xlPageBreakNone** . Read/write **Long** .|
|[Parent](range-parent-property-excel.md)|Returns the parent object for the specified object. Read-only.|
|[Phonetic](range-phonetic-property-excel.md)|Returns the  **[Phonetic](phonetic-object-excel.md)** object, which contains information about a specific phonetic text string in a cell.|
|[Phonetics](range-phonetics-property-excel.md)|Returns the  **[Phonetics](phonetics-object-excel.md)** collection of the range. Read only.|
|[PivotCell](range-pivotcell-property-excel.md)|Returns a  **[PivotCell](pivotcell-object-excel.md)** object that represents a cell in a PivotTable report.|
|[PivotField](range-pivotfield-property-excel.md)|Returns a  **[PivotField](pivotfield-object-excel.md)** object that represents the PivotTable field containing the upper-left corner of the specified range.|
|[PivotItem](range-pivotitem-property-excel.md)|Returns a  **[PivotItem](pivotitem-object-excel.md)** object that represents the PivotTable item containing the upper-left corner of the specified range.|
|[PivotTable](range-pivottable-property-excel.md)|Returns a  **[PivotTable](pivottable-object-excel.md)** object that represents the PivotTable report containing the upper-left corner of the specified range.|
|[Precedents](range-precedents-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents all the precedents of a cell. This can be a multiple selection (a union of **Range** objects) if there's more than one precedent. Read-only.|
|[PrefixCharacter](range-prefixcharacter-property-excel.md)|Returns the prefix character for the cell. Read-only  **Variant** .|
|[Previous](range-previous-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents the next cell.|
|[QueryTable](range-querytable-property-excel.md)|Returns a  **[QueryTable](querytable-object-excel.md)** object that represents the query table that intersects the specified **[Range](range-object-excel.md)** object.|
|[Range](range-range-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents a cell or a range of cells.|
|[ReadingOrder](range-readingorder-property-excel.md)|Returns or sets the reading order for the specified object. Can be one of the following constants:  **xlRTL** (right-to-left), **xlLTR** (left-to-right), or **xlContext** . Read/write **Long** .|
|[Resize](range-resize-property-excel.md)|Resizes the specified range. Returns a  **[Range](range-object-excel.md)** object that represents the resized range.|
|[Row](range-row-property-excel.md)|Returns the number of the first row of the first area in the range. Read-only  **Long** .|
|[RowHeight](range-rowheight-property-excel.md)|Returns or sets the height of the first row in the range specified, measured in points. Read/write  **Variant** .|
|[Rows](range-rows-property-excel.md)|Returns a  **[Range](range-object-excel.md)** object that represents the rows in the specified range. Read-only **Range** object.|
|[ServerActions](range-serveractions-property-excel.md)|Specifies the actions that can be performed on the SharePoint server for a  **Range** object.|
|[ShowDetail](range-showdetail-property-excel.md)| **True** if the outline is expanded for the specified range (so that the detail of the column or row is visible). The specified range must be a single summary column or row in an outline. Read/write **Variant** . For the **PivotItem** object (or the **Range** object if the range is in a PivotTable report), this property is set to **True** if the item is showing detail.|
|[ShrinkToFit](range-shrinktofit-property-excel.md)|Returns or sets a  **Variant** value that indicates if text automatically shrinks to fit in the available column width.|
|[SoundNote](range-soundnote-property-excel.md)|This property should not be used. Sound notes have been removed from Microsoft Excel.|
|[SparklineGroups](range-sparklinegroups-property-excel.md)|Returns a  **[SparklineGroups](sparklinegroups-object-excel.md)** object that represents an existing group of sparklines from the specified range. Read-only|
|[Style](range-style-property-excel.md)|Returns or sets a  **Variant** value, containing a **[Style](style-object-excel.md)** object, that represents the style of the specified range.|
|[Summary](range-summary-property-excel.md)| **True** if the range is an outlining summary row or column. The range should be a row or a column. Read-only **Variant** .|
|[Text](range-text-property-excel.md)|Returns or sets the text for the specified object. Read-only  **String** .|
|[Top](range-top-property-excel.md)|Returns a  **Variant** value that represents the distance, in points, from the top edge of row 1 to the top edge of the range.|
|[UseStandardHeight](range-usestandardheight-property-excel.md)| **True** if the row height of the **Range** object equals the standard height of the sheet. Returns **Null** if the range contains more than one row and the rows aren't all the same height. Read/write **Variant** .|
|[UseStandardWidth](range-usestandardwidth-property-excel.md)| **True** if the column width of the **Range** object equals the standard width of the sheet. Returns **null** if the range contains more than one column and the columns aren't all the same width. Read/write **Variant** .|
|[Validation](range-validation-property-excel.md)|Returns the  **[Validation](validation-object-excel.md)** object that represents data validation for the specified range. Read-only.|
|[Value](range-value-property-excel.md)|Returns or sets a  **Variant** value that represents the value of the specified range.|
|[Value2](range-value2-property-excel.md)|Returns or sets the cell value. Read/write  **Variant** .|
|[VerticalAlignment](range-verticalalignment-property-excel.md)|Returns or sets a  **Variant** value that represents the vertical alignment of the specified object.|
|[Width](range-width-property-excel.md)|Returns a  **Variant** value that represents the width, in units, of the range.|
|[Worksheet](range-worksheet-property-excel.md)|Returns a  **[Worksheet](worksheet-object-excel.md)** object that represents the worksheet containing the specified range. Read-only.|
|[WrapText](range-wraptext-property-excel.md)|Returns or sets a  **Variant** value that indicates if Microsoft Excel wraps the text in the object.|
|[XPath](range-xpath-property-excel.md)|Returns an  **[XPath](xpath-object-excel.md)** object that represents the Xpath of the element mapped to the specified **[Range](range-object-excel.md)** object. The context of the range determines whether or not the action succeeds or returns an empty object. Read-only.|

