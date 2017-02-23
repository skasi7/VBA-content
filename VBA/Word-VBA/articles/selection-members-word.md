---
title: Selection Members (Word)
ms.prod: WORD
ms.assetid: 71e67a43-d40a-ad9a-8ef2-c5c487733e0d
---


# Selection Members (Word)
Represents the current selection in a window or pane. A selection represents either a selected (or highlighted) area in the document, or it represents the insertion point if nothing in the document is selected. There can be only one  **Selection** object per document window pane, and only one **Selection** object in the entire application can be active.

Represents the current selection in a window or pane. A selection represents either a selected (or highlighted) area in the document, or it represents the insertion point if nothing in the document is selected. There can be only one  **Selection** object per document window pane, and only one **Selection** object in the entire application can be active.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[BoldRun](selection-boldrun-method-word.md)|Adds the bold character format to or removes it from the current run.|
|[Calculate](selection-calculate-method-word.md)|Calculates a mathematical expression within a selection. Returns the result as a  **Single** .|
|[ClearCharacterAllFormatting](selection-clearcharacterallformatting-method-word.md)|Removes all character formatting (formatting applied either through character styles or manually applied formatting) from the selected text.|
|[ClearCharacterDirectFormatting](selection-clearcharacterdirectformatting-method-word.md)|Removes character formatting (formatting that has been applied manually using the buttons on the ribbon or through the dialog boxes) from the selected text.|
|[ClearCharacterStyle](selection-clearcharacterstyle-method-word.md)|Removes character formatting that has been applied through character styles from the selected text.|
|[ClearFormatting](selection-clearformatting-method-word.md)|Removes text and paragraph formatting from a selection.|
|[ClearParagraphAllFormatting](selection-clearparagraphallformatting-method-word.md)|Removes all paragraph formatting (formatting applied either through paragraph styles or manually applied formatting) from the selected text.|
|[ClearParagraphDirectFormatting](selection-clearparagraphdirectformatting-method-word.md)|Removes paragraph formatting that has been applied manually (using the buttons on the ribbon or through the dialog boxes) from the selected text.|
|[ClearParagraphStyle](selection-clearparagraphstyle-method-word.md)|Removes paragraph formatting that has been applied through paragraph styles from the selected text.|
|[Collapse](selection-collapse-method-word.md)|Collapses a selection to the starting or ending position. After a selection is collapsed, the starting and ending points are equal.|
|[ConvertToTable](selection-converttotable-method-word.md)|Converts text within a range to a table. Returns the table as a  **Table** object.|
|[Copy](selection-copy-method-word.md)|Copies the specified selection to the Clipboard.|
|[CopyAsPicture](selection-copyaspicture-method-word.md)|The  **CopyAsPicture** method works the same way as the **Copy** method.|
|[CopyFormat](selection-copyformat-method-word.md)|Copies the character formatting of the first character in the selected text.|
|[CreateAutoTextEntry](selection-createautotextentry-method-word.md)|Adds a new  **[AutoTextEntry](autotextentry-object-word.md)** object to the **[AutoTextEntries](autotextentries-object-word.md)** collection, based on the current selection.|
|[CreateTextbox](selection-createtextbox-method-word.md)|Adds a default-size text box around the selection.|
|[Cut](selection-cut-method-word.md)|Removes the specified object from the document and moves it to the Clipboard.|
|[Delete](selection-delete-method-word.md)|Deletes the specified number of characters or words.|
|[DetectLanguage](selection-detectlanguage-method-word.md)|Analyzes the specified text to determine the language that it is written in.|
|[EndKey](selection-endkey-method-word.md)|Moves or extends the selection to the end of the specified unit.|
|[EndOf](selection-endof-method-word.md)|Moves or extends the ending character position of a range or selection to the end of the nearest specified text unit.|
|[EscapeKey](selection-escapekey-method-word.md)|Cancels a mode such as extend or column select (equivalent to pressing the ESC key).|
|[Expand](selection-expand-method-word.md)|Expands the specified range or selection. Returns the number of characters added to the range or selection.  **Long** .|
|[ExportAsFixedFormat](selection-exportasfixedformat-method-word.md)|Saves the current selection as PDF or XPS format. .|
|[Extend](selection-extend-method-word.md)|Turns on extend mode, or if extend mode is already on, extends the selection to the next larger unit of text.|
|[GoTo](selection-goto-method-word.md)|Moves the insertion point to the character position immediately preceding the specified item, and returns a  **Range** object (except for the **wdGoToGrammaticalError** , **wdGoToProofreadingError** , or **wdGoToSpellingError** constant).|
|[GoToEditableRange](selection-gotoeditablerange-method-word.md)|Returns a  **Range** object that represents an area of a document that can be modified by the specified user or group of users.|
|[GoToNext](selection-gotonext-method-word.md)|Returns a  **Range** object that refers to the start position of the next item or location specified by the What argument. If you apply this method to the **Selection** object, the method moves the selection to the specified item (except for the **wdGoToGrammaticalError** , **wdGoToProofreadingError** , and **wdGoToSpellingError** constants).|
|[GoToPrevious](selection-gotoprevious-method-word.md)|Returns a  **Range** object that refers to the start position of the previous item or location specified by the What argument. If applied to a **Selection** object, **GoToPrevious** moves the selection to the specified item. **Range** object.|
|[HomeKey](selection-homekey-method-word.md)|Moves or extends the selection to the beginning of the specified unit. This method returns an integer that indicates the number of characters the selection was actually moved, or it returns 0 (zero) if the move was unsuccessful.This method corresponds to functionality of the HOME key.|
|[InRange](selection-inrange-method-word.md)| **True** if the selection to which the method is applied is contained within the range specified by the Range argument.|
|[InsertAfter](selection-insertafter-method-word.md)|Inserts the specified text at the end of a range or selection.|
|[InsertBefore](selection-insertbefore-method-word.md)|Inserts the specified text before the specified selection. .|
|[InsertBreak](selection-insertbreak-method-word.md)|Inserts a page, column, or section break.|
|[InsertCaption](selection-insertcaption-method-word.md)|Inserts a caption immediately preceding or following the specified selection.|
|[InsertCells](selection-insertcells-method-word.md)|Adds cells to an existing table.|
|[InsertColumns](selection-insertcolumns-method-word.md)|Inserts columns to the left of the column that contains the selection.|
|[InsertColumnsRight](selection-insertcolumnsright-method-word.md)|Inserts columns to the right of the current selection.|
|[InsertCrossReference](selection-insertcrossreference-method-word.md)|Inserts a cross-reference to a heading, bookmark, footnote, or endnote, or to an item for which a caption label is defined (for example, an equation, figure, or table).|
|[InsertDateTime](selection-insertdatetime-method-word.md)|Inserts the current date or time, or both, either as text or as a TIME field.|
|[InsertFile](selection-insertfile-method-word.md)|Inserts all or part of the specified file.|
|[InsertFormula](selection-insertformula-method-word.md)|Inserts an = (Formula) field that contains a formula at the selection.|
|[InsertNewPage](selection-insertnewpage-method-word.md)|Inserts a new page at the position of the Insertion Point.|
|[InsertParagraph](selection-insertparagraph-method-word.md)|Replaces the specified selection with a new paragraph.|
|[InsertParagraphAfter](selection-insertparagraphafter-method-word.md)|Inserts a paragraph mark after a selection.|
|[InsertParagraphBefore](selection-insertparagraphbefore-method-word.md)|Inserts a new paragraph before the specified selection or range.|
|[InsertRows](selection-insertrows-method-word.md)|Inserts the specified number of new rows above the row that contains the selection. If the selection isn't in a table, an error occurs.|
|[InsertRowsAbove](selection-insertrowsabove-method-word.md)|Inserts rows above the current selection.|
|[InsertRowsBelow](selection-insertrowsbelow-method-word.md)|Inserts rows below the current selection.|
|[InsertStyleSeparator](selection-insertstyleseparator-method-word.md)|Inserts a special hidden paragraph mark that allows Microsoft Word to join paragraphs formatted using different paragraph styles, so lead-in headings can be inserted into a table of contents.|
|[InsertSymbol](selection-insertsymbol-method-word.md)|Inserts a symbol in place of the specified selection.|
|[InsertXML](selection-insertxml-method-word.md)|Inserts the specified XML into the document at the cursor, replacing any selected text.|
|[InStory](selection-instory-method-word.md)| **True** if the selection to which this method is applied is in the same story as the range specified by the Range argument.|
|[IsEqual](selection-isequal-method-word.md)| **True** if the selection to which this method is applied is equal to the range specified by the Range argument.|
|[ItalicRun](selection-italicrun-method-word.md)|Adds the italic character format to or removes it from the current run.|
|[LtrPara](selection-ltrpara-method-word.md)|Sets the reading order and alignment of the specified paragraphs to left-to-right.|
|[LtrRun](selection-ltrrun-method-word.md)|Sets the reading order and alignment of the specified run to left-to-right.|
|[Move](selection-move-method-word.md)|Collapses the specified selection to its start or end position and then moves the collapsed object by the specified number of units. This method returns a  **Long** value that represents the number of units by which the selection was moved, or it returns 0 (zero) if the move was unsuccessful.|
|[MoveDown](selection-movedown-method-word.md)|Moves the selection down and returns the number of units it has been moved.|
|[MoveEnd](selection-moveend-method-word.md)|Moves the ending character position of a range or selection.|
|[MoveEndUntil](selection-moveenduntil-method-word.md)|Moves the end position of the specified selection until any of the specified characters are found in the document.|
|[MoveEndWhile](selection-moveendwhile-method-word.md)|Moves the ending character position of a selection while any of the specified characters are found in the document.|
|[MoveLeft](selection-moveleft-method-word.md)|Moves the selection to the left and returns the number of units it has been moved.|
|[MoveRight](selection-moveright-method-word.md)|Moves the selection to the right and returns the number of units it has been moved.|
|[MoveStart](selection-movestart-method-word.md)|Moves the start position of the specified selection.|
|[MoveStartUntil](selection-movestartuntil-method-word.md)|Moves the start position of the specified selection until one of the specified characters is found in the document. If the movement is backward through the document, the selection is expanded.|
|[MoveStartWhile](selection-movestartwhile-method-word.md)|Moves the start position of the specified selection while any of the specified characters are found in the document.|
|[MoveUntil](selection-moveuntil-method-word.md)|Moves the specified selection until one of the specified characters is found in the document.|
|[MoveUp](selection-moveup-method-word.md)|Moves the selection up and returns the number of units that it has been moved.|
|[MoveWhile](selection-movewhile-method-word.md)|Moves the specified selection while any of the specified characters are found in the document.|
|[Next](selection-next-method-word.md)|Returns a  **Range** object that represents the next unit relative to the specified selection.|
|[NextField](selection-nextfield-method-word.md)|Selects the next field.|
|[NextRevision](selection-nextrevision-method-word.md)|Locates and returns the next tracked change as a  **Revision** object.|
|[NextSubdocument](selection-nextsubdocument-method-word.md)|Moves the selection to the next subdocument.|
|[Paste](selection-paste-method-word.md)|Inserts the contents of the Clipboard at the specified selection.|
|[PasteAndFormat](selection-pasteandformat-method-word.md)|Pastes the selected table cells and formats them as specified.|
|[PasteAppendTable](selection-pasteappendtable-method-word.md)|Merges pasted cells into an existing table by inserting the pasted rows between the selected rows. No cells are overwritten.|
|[PasteAsNestedTable](selection-pasteasnestedtable-method-word.md)|Pastes a cell or group of cells as a nested table into the selection.|
|[PasteExcelTable](selection-pasteexceltable-method-word.md)|Pastes and formats a Microsoft Excel table.|
|[PasteFormat](selection-pasteformat-method-word.md)|Applies formatting copied with the  **CopyFormat** method to the selection.|
|[PasteSpecial](selection-pastespecial-method-word.md)|Inserts the contents of the Clipboard.|
|[Previous](selection-previous-method-word.md)|Moves the selected text by the specified number of units, and returns a  **Range** object relative to the collapsed selection.|
|[PreviousField](selection-previousfield-method-word.md)|Selects and returns the previous field.|
|[PreviousRevision](selection-previousrevision-method-word.md)|Locates and returns the previous tracked change as a  **Revision** object.|
|[PreviousSubdocument](selection-previoussubdocument-method-word.md)|Moves the selection to the previous subdocument.|
|[ReadingModeGrowFont](selection-readingmodegrowfont-method-word.md)|Increases the size of the displayed text one point size when the document is displayed in Reading mode.|
|[ReadingModeShrinkFont](selection-readingmodeshrinkfont-method-word.md)|Decreases the size of the displayed text one point size when the document is displayed in Reading mode.|
|[RtlPara](selection-rtlpara-method-word.md)|Sets the reading order and alignment of the specified paragraphs to right-to-left.|
|[RtlRun](selection-rtlrun-method-word.md)|Sets the reading order and alignment of the specified run to right-to-left.|
|[Select](selection-select-method-word.md)|Selects the specified text.|
|[SelectCell](selection-selectcell-method-word.md)|Selects the entire cell containing the current selection.|
|[SelectColumn](selection-selectcolumn-method-word.md)|Selects the column that contains the insertion point, or selects all columns that contain the selection.|
|[SelectCurrentAlignment](selection-selectcurrentalignment-method-word.md)|Extends the selection forward until text with a different paragraph alignment is encountered.|
|[SelectCurrentColor](selection-selectcurrentcolor-method-word.md)|Extends the selection forward until text with a different color is encountered.|
|[SelectCurrentFont](selection-selectcurrentfont-method-word.md)|Extends the selection forward until text in a different font or font size is encountered.|
|[SelectCurrentIndent](selection-selectcurrentindent-method-word.md)|Extends the selection forward until text with different left or right paragraph indents is encountered.|
|[SelectCurrentSpacing](selection-selectcurrentspacing-method-word.md)|Extends the selection forward until a paragraph with different line spacing is encountered.|
|[SelectCurrentTabs](selection-selectcurrenttabs-method-word.md)|Extends the selection forward until a paragraph with different tab stops is encountered.|
|[SelectRow](selection-selectrow-method-word.md)|Selects the row that contains the insertion point, or selects all rows that contain the selection.|
|[SetRange](selection-setrange-method-word.md)|Sets the starting and ending character positions for the selection.|
|[Shrink](selection-shrink-method-word.md)|Shrinks the selection to the next smaller unit of text.|
|[ShrinkDiscontiguousSelection](selection-shrinkdiscontiguousselection-method-word.md)|Cancels the selection of all but the most recently selected text when a selection contains multiple, unconnected selections.|
|[Sort](selection-sort-method-word.md)|Sorts the paragraphs in the specified selection.|
|[SortAscending](selection-sortascending-method-word.md)|Sorts paragraphs or table rows in ascending alphanumeric order.|
|[SortByHeadings](selection-sortbyheadings-method-word.md)|Sorts the headings in the specified selection.|
|[SortDescending](selection-sortdescending-method-word.md)|Sorts paragraphs or table rows within the selection in descending alphanumeric order.|
|[SplitTable](selection-splittable-method-word.md)|Inserts an empty paragraph above the first row in the selection. .|
|[StartOf](selection-startof-method-word.md)|Moves or extends the start position of the specified range or selection to the beginning of the nearest specified text unit. This method returns a  **Long** that indicates the number of characters by which the range or selection was moved or extended. The method returns a negative number if the movement is backward through the document.|
|[ToggleCharacterCode](selection-togglecharactercode-method-word.md)|Switches a selection between a Unicode character and its corresponding hexadecimal value.|
|[TypeBackspace](selection-typebackspace-method-word.md)|Deletes the character preceding a collapsed selection (an insertion point).|
|[TypeParagraph](selection-typeparagraph-method-word.md)|Inserts a new, blank paragraph.|
|[TypeText](selection-typetext-method-word.md)|Inserts the specified text.|
|[WholeStory](selection-wholestory-method-word.md)|Expands a selection to include the entire story.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Active](selection-active-property-word.md)| **True** if the selection in the specified window or pane is active. Read-only **Boolean** .|
|[Application](selection-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[BookmarkID](selection-bookmarkid-property-word.md)|Returns the number of the bookmark that encloses the beginning of the specified selection. Read-only  **Long** .|
|[Bookmarks](selection-bookmarks-property-word.md)|Returns a  **[Bookmarks](bookmarks-object-word.md)** collection that represents all the bookmarks in a document, range, or selection. Read-only.|
|[Borders](selection-borders-property-word.md)|Returns a  **[Borders](borders-object-word.md)** collection that represents all the borders for the specified object.|
|[Cells](selection-cells-property-word.md)|Returns a  **[Cells](cells-object-word.md)** collection that represents the table cells in a selection. Read-only.|
|[Characters](selection-characters-property-word.md)|Returns a  **[Characters](characters-object-word.md)** collection that represents the characters in a document, range, or selection. Read-only.|
|[ChildShapeRange](selection-childshaperange-property-word.md)|Returns a  **[ShapeRange](shaperange-object-word.md)** collection representing the child shapes contained within a selection.|
|[Columns](selection-columns-property-word.md)|Returns a  **Columns** collection that represents all the table columns in a selection. Read-only.|
|[ColumnSelectMode](selection-columnselectmode-property-word.md)| **True** if column selection mode is active. Read/write **Boolean** .|
|[Comments](selection-comments-property-word.md)|Returns a  **[Comments](comments-object-word.md)** collection that represents all the comments in the specified. Read-only.|
|[Creator](selection-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[Document](selection-document-property-word.md)|Returns a  **[Document](document-object-word.md)** object associated with the specified selection. Read-only.|
|[Editors](selection-editors-property-word.md)|Returns an  **[Editors](editors-object-word.md)** object that represents all the users authorized to modify a selection within a document.|
|[End](selection-end-property-word.md)|Returns or sets the ending character position of a selection. Read/write  **Long** .|
|[EndnoteOptions](selection-endnoteoptions-property-word.md)|Returns an  **[EndnoteOptions](endnoteoptions-object-word.md)** object that represents the endnotes in a selection.|
|[Endnotes](selection-endnotes-property-word.md)|Returns an  **[Endnotes](endnotes-object-word.md)** collection that represents all the endnotes conatined within a selection. Read-only.|
|[EnhMetaFileBits](selection-enhmetafilebits-property-word.md)|Returns a  **Variant** that represents a picture representation of how a selection or range of text appears.|
|[ExtendMode](selection-extendmode-property-word.md)| **True** if Extend mode is active. Read/write **Boolean** .|
|[Fields](selection-fields-property-word.md)|Returns a read-only  **[Fields](fields-object-word.md)** collection that represents all the fields in the selection.|
|[Find](selection-find-property-word.md)|Returns a  **[Find](find-object-word.md)** object that contains the criteria for a find operation. Read-only.|
|[FitTextWidth](selection-fittextwidth-property-word.md)|Returns or sets the width (in the current measurement units) in which Microsoft Word fits the text in the current selection. Read/write  **Single** .|
|[Flags](selection-flags-property-word.md)|Returns or sets properties of the selection. Read/write  **WdSelectionFlags** .|
|[Font](selection-font-property-word.md)|Returns or sets a  **[Font](font-object-word.md)** object that represents the character formatting of the specified object. Read/write.|
|[FootnoteOptions](selection-footnoteoptions-property-word.md)|Returns  **[FootnoteOptions](footnoteoptions-object-word.md)** object that represents the footnotes in a selection.|
|[Footnotes](selection-footnotes-property-word.md)|Returns a  **[Footnotes](footnotes-object-word.md)** collection that represents all the footnotes in a range, selection, or document. Read-only.|
|[FormattedText](selection-formattedtext-property-word.md)|Returns or sets a  **[Range](range-object-word.md)** object that includes the formatted text in the specified range or selection. Read/write.|
|[FormFields](selection-formfields-property-word.md)|Returns a  **[FormFields](formfields-object-word.md)** collection that represents all the form fields in the selection. Read-only.|
|[Frames](selection-frames-property-word.md)|Returns a  **[Frames](frames-object-word.md)** collection that represents all the frames in a selection. Read-only.|
|[HasChildShapeRange](selection-haschildshaperange-property-word.md)| **True** if the selection contains child shapes. Read-only **Boolean** .|
|[HeaderFooter](selection-headerfooter-property-word.md)|Returns a  **[HeaderFooter](headerfooter-object-word.md)** object for the specified selection. Read-only.|
|[HTMLDivisions](selection-htmldivisions-property-word.md)|Returns an  **[HTMLDivisions](htmldivisions-object-word.md)** object that represents an HTML division in a Web document.|
|[Hyperlinks](selection-hyperlinks-property-word.md)|Returns a  **[Hyperlinks](hyperlinks-object-word.md)** collection that represents all the hyperlinks in the specified selection. Read-only.|
|[Information](selection-information-property-word.md)|Returns information about the specified selection. Read-only  **Variant** .|
|[InlineShapes](selection-inlineshapes-property-word.md)|Returns an  **[InlineShapes](inlineshapes-object-word.md)** collection that represents all the **InlineShape** objects in a selection. Read-only.|
|[IPAtEndOfLine](selection-ipatendofline-property-word.md)| **True** if the insertion point is at the end of a line that wraps to the next line. Read-only **Boolean** .|
|[IsEndOfRowMark](selection-isendofrowmark-property-word.md)| **True** if the specified selection or range is collapsed and is located at the end-of-row mark in a table. Read-only **Boolean** .|
|[LanguageDetected](selection-languagedetected-property-word.md)|Returns or sets a  **Boolean** that specifies whether Microsoft Word has detected the language of the selected text.|
|[LanguageID](selection-languageid-property-word.md)|Returns or sets the language for the specified object. Read/write .|
|[LanguageIDFarEast](selection-languageidfareast-property-word.md)|Returns or sets an East Asian language for the specified object. Read/write  **WdLanguageID** .|
|[LanguageIDOther](selection-languageidother-property-word.md)|Returns or sets the language for the specified object. Read/write  **WdLanguageID** .|
|[NoProofing](selection-noproofing-property-word.md)| **True** if the spelling and grammar checker ignores the specified text. Returns **wdUndefined** if the **NoProofing** property is set to **True** for only some of the specified text. Read/write **Long** .|
|[OMaths](selection-omaths-property-word.md)|Returns an  **[OMaths](omaths-object-word.md)** collection that represents the **[OMath](omath-object-word.md)** objects within the current selection. Read-only.|
|[Orientation](selection-orientation-property-word.md)|Returns or sets the orientation of text in a selection when the Text Direction feature is enabled. Read/write  **WdTextOrientation** .|
|[PageSetup](selection-pagesetup-property-word.md)|Returns a  **[PageSetup](pagesetup-object-word.md)** object that's associated with the specified selection.|
|[ParagraphFormat](selection-paragraphformat-property-word.md)|Returns or sets a  **[ParagraphFormat](paragraphformat-object-word.md)** object that represents the paragraph settings for the specified selection. Read/write.|
|[Paragraphs](selection-paragraphs-property-word.md)|Returns a  **[Paragraphs](paragraphs-object-word.md)** collection that represents all the paragraphs in the specified selection. Read-only.|
|[Parent](selection-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **[Selection](selection-object-word.md)** object.|
|[PreviousBookmarkID](selection-previousbookmarkid-property-word.md)|Returns the number of the last bookmark that starts before or at the same place as the specified selection or range; returns 0 (zero) if there is no corresponding bookmark. Read-only  **Long** .|
|[Range](selection-range-property-word.md)|Returns a  **[Range](range-object-word.md)** object that represents the portion of a document that's contained in the specified object.|
|[Rows](selection-rows-property-word.md)|Returns a  **[Rows](rows-object-word.md)** collection that represents all the table rows in a range, selection, or table. Read-only.|
|[Sections](selection-sections-property-word.md)|Returns a  **[Sections](sections-object-word.md)** collection that represents the sections in the specified selection. Read-only.|
|[Sentences](selection-sentences-property-word.md)|Returns a  **[Sentences](sentences-object-word.md)** collection that represents all the sentences in the selection. Read-only.|
|[Shading](selection-shading-property-word.md)|Returns a  **[Shading](shading-object-word.md)** object that refers to the shading formatting for the specified selection.|
|[ShapeRange](selection-shaperange-property-word.md)|Returns a  **[ShapeRange](shaperange-object-word.md)** collection that represents all the **Shape** objects in the selection. Read-only.|
|[Start](selection-start-property-word.md)|Returns or sets the starting character position of a selection. Read/write  **Long** .|
|[StartIsActive](selection-startisactive-property-word.md)| **True** if the beginning of the selection is active. Read/write **Boolean** .|
|[StoryLength](selection-storylength-property-word.md)|Returns the number of characters in the story that contains the specified selection. Read-only  **Long** .|
|[StoryType](selection-storytype-property-word.md)|Returns the story type for the specified selection. Read-only  **WdStoryType** .|
|[Style](selection-style-property-word.md)|Returns or sets the style for the specified object. To set this property, specify the local name of the style, an integer, a  **WdBuiltinStyle** constant, or an object that represents the style. For a list of valid constants, consult the Microsoft Visual Basic Object Browser. Read/write **Variant** .|
|[Tables](selection-tables-property-word.md)|Returns a  **[Tables](tables-object-word.md)** collection that represents all the tables in the specified selection. Read-only.|
|[Text](selection-text-property-word.md)|Returns or sets the text in the specified selection. Read/write  **String** .|
|[TopLevelTables](selection-topleveltables-property-word.md)|Returns a  **[Tables](tables-object-word.md)** collection that represents the tables at the outermost nesting level in the current selection. Read-only.|
|[Type](selection-type-property-word.md)|Returns the selection type. Read-only  **[WdSelectionType](wdselectiontype-enumeration-word.md)** .|
|[WordOpenXML](selection-wordopenxml-property-word.md)|Returns a  **String** that represents the XML contained within the selection in the Microsoft Word Open XML format. Read-only.|
|[Words](selection-words-property-word.md)|Returns a  **[Words](words-object-word.md)** collection that represents all the words in a selection. Read-only.|
|[XML](selection-xml-property-word.md)|Returns a  **String** that represents the XML text in the specified object. .|

