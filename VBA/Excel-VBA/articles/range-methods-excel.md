---
title: Range Methods (Excel)
ms.prod: EXCEL
ms.assetid: 264d41e4-1616-49ea-84f1-06837d656e51
---


# Range Methods (Excel)

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

