---
title: Range Properties (Excel)
ms.prod: EXCEL
ms.assetid: 76025eb0-8ed1-4f98-b944-e60e655ae24c
---


# Range Properties (Excel)

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

