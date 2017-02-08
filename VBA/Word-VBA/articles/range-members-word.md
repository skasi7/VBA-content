---
title: Range Members (Word)
ms.prod: WORD
ms.assetid: 3c4a36d9-2a80-5aaf-827b-275a52bfa193
---


# Range Members (Word)
Represents a contiguous area in a document. Each  **Range** object is defined by a starting and ending character position.

Represents a contiguous area in a document. Each  **Range** object is defined by a starting and ending character position.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[AutoFormat](range-autoformat-method-word.md)|Automatically formats a document. Use the  **Kind** property to specify a document type.|
|[Calculate](range-calculate-method-word.md)|Calculates a mathematical expression within a range or selection. Returns the result as a  **Single** .|
|[CheckGrammar](range-checkgrammar-method-word.md)|Begins a spelling and grammar check for the specified range.|
|[CheckSpelling](range-checkspelling-method-word.md)|Begins a spelling check for the specified document or range.|
|[CheckSynonyms](range-checksynonyms-method-word.md)|Displays the  **Thesaurus** dialog box, which lists alternative word choices, or synonyms, for the text in the specified range.|
|[Collapse](range-collapse-method-word.md)|Collapses a range or selection to the starting or ending position. After a range or selection is collapsed, the starting and ending points are equal.|
|[ComputeStatistics](range-computestatistics-method-word.md)|Returns a  **Long** that represents a statistic based on the contents of the specified range.|
|[ConvertHangulAndHanja](range-converthangulandhanja-method-word.md)|Converts the specified range from hangul to hanja or vice versa.|
|[ConvertToTable](range-converttotable-method-word.md)|Converts text within a range to a table. Returns the table as a  **Table** object.|
|[Copy](range-copy-method-word.md)|Copies the specified range to the Clipboard.|
|[CopyAsPicture](range-copyaspicture-method-word.md)|The  **CopyAsPicture** method works the same way as the **Copy** method.|
|[Cut](range-cut-method-word.md)|Removes the specified object from the document and places it on the Clipboard.|
|[Delete](range-delete-method-word.md)|Deletes the specified number of characters or words.|
|[DetectLanguage](range-detectlanguage-method-word.md)|Analyzes the specified text to determine the language that it is written in.|
|[EndOf](range-endof-method-word.md)|Moves or extends the ending character position of a range to the end of the nearest specified text unit.|
|[Expand](range-expand-method-word.md)|Expands the specified range or selection. Returns the number of characters added to the range or selection.  **Long** .|
|[ExportAsFixedFormat](range-exportasfixedformat-method-word.md)|Saves a portion of a document as PDF or XPS format.|
|[ExportFragment](range-exportfragment-method-word.md)| Exports the selected range into a document for use as a document fragment.|
|[GetSpellingSuggestions](range-getspellingsuggestions-method-word.md)|Returns a  **SpellingSuggestions** collection that represents the words suggested as spelling replacements for the first word in the specified range.|
|[GoTo](range-goto-method-word.md)|Returns a  **Range** object that represents the start position of the specified item, such as a page, bookmark, or field.|
|[GoToEditableRange](range-gotoeditablerange-method-word.md)|Returns a  **Range** object that represents an area of a document that can be modified by the specified user or group of users.|
|[GoToNext](range-gotonext-method-word.md)|Returns a  **Range** object that refers to the start position of the next item or location specified by the What argument. .|
|[GoToPrevious](range-gotoprevious-method-word.md)|Returns a  **Range** object that refers to the start position of the previous item or location specified by the What argument.|
|[ImportFragment](range-importfragment-method-word.md)|Imports a document fragment into the document at the specified range.|
|[InRange](range-inrange-method-word.md)|Returns  **True** if the range to which the method is applied is contained in the range specified by the Range argument.|
|[InsertAfter](range-insertafter-method-word.md)|Inserts the specified text at the end of a range.|
|[InsertAlignmentTab](range-insertalignmenttab-method-word.md)|Inserts an absolute tab that is always positioned in the same spot, relative to either the margins or indents.|
|[InsertAutoText](range-insertautotext-method-word.md)|Attempts to match the text in the specified range or the text surrounding the range with an existing AutoText entry name.|
|[InsertBefore](range-insertbefore-method-word.md)|Inserts the specified text before the specified range.|
|[InsertBreak](range-insertbreak-method-word.md)|Inserts a page, column, or section break.|
|[InsertCaption](range-insertcaption-method-word.md)|Inserts a caption immediately preceding or following the specified range.|
|[InsertCrossReference](range-insertcrossreference-method-word.md)|Inserts a cross-reference to a heading, bookmark, footnote, or endnote, or to an item for which a caption label is defined (for example, an equation, figure, or table).|
|[InsertDatabase](range-insertdatabase-method-word.md)|Retrieves data from a data source (for example, a separate Microsoft Word document, a Microsoft Excel worksheet, or a Microsoft Access database) and inserts the data as a table in place of the specified range.|
|[InsertDateTime](range-insertdatetime-method-word.md)|Inserts the current date or time, or both, either as text or as a TIME field.|
|[InsertFile](range-insertfile-method-word.md)|Inserts all or part of the specified file.|
|[InsertParagraph](range-insertparagraph-method-word.md)|Replaces the specified range with a new paragraph.|
|[InsertParagraphAfter](range-insertparagraphafter-method-word.md)|Inserts a paragraph mark after a range.|
|[InsertParagraphBefore](range-insertparagraphbefore-method-word.md)|Inserts a new paragraph before the specified range.|
|[InsertSymbol](range-insertsymbol-method-word.md)|Inserts a symbol in place of the specified range.|
|[InsertXML](range-insertxml-method-word.md)|Inserts the specified XML into the document at the specified range, replacing any text contained within the range.|
|[InStory](range-instory-method-word.md)| **True** if the range to which this method is applied is in the same story as the range specified by the Range argument.|
|[IsEqual](range-isequal-method-word.md)| **True** if the range to which this method is applied is equal to the range specified by the Range argument.|
|[LookupNameProperties](range-lookupnameproperties-method-word.md)|Looks up a name in the global address book list and displays the  **Properties** dialog box, which includes information about the specified name.|
|[ModifyEnclosure](range-modifyenclosure-method-word.md)|Adds, modifies, or removes an enclosure around the specified character or characters.|
|[Move](range-move-method-word.md)|Collapses the specified range to its start or end position and then moves the collapsed object by the specified number of units.|
|[MoveEnd](range-moveend-method-word.md)|Moves the ending character position of a range. .|
|[MoveEndUntil](range-moveenduntil-method-word.md)|Moves the end position of the specified range until any of the specified characters are found in the document. If the movement is forward in the document, the range is expanded.|
|[MoveEndWhile](range-moveendwhile-method-word.md)|Moves the ending character position of a range while any of the specified characters are found in the document.|
|[MoveStart](range-movestart-method-word.md)|Moves the start position of the specified range.|
|[MoveStartUntil](range-movestartuntil-method-word.md)|Moves the start position of the specified range until one of the specified characters is found in the document.|
|[MoveStartWhile](range-movestartwhile-method-word.md)|Moves the start position of the specified range while any of the specified characters are found in the document.|
|[MoveUntil](range-moveuntil-method-word.md)|Moves the specified range until one of the specified characters is found in the document.|
|[MoveWhile](range-movewhile-method-word.md)|Moves the specified range while any of the specified characters are found in the document.|
|[Next](range-next-method-word.md)|Returns a  **Range** object that represents the specified unit relative to the specified range.|
|[NextSubdocument](range-nextsubdocument-method-word.md)|Moves the range to the next subdocument.|
|[Paste](range-paste-method-word.md)|Inserts the contents of the Clipboard at the specified range.|
|[PasteAndFormat](range-pasteandformat-method-word.md)|Pastes the selected table cells and formats them as specified.|
|[PasteAppendTable](range-pasteappendtable-method-word.md)|Merges pasted cells into an existing table by inserting the pasted rows between the selected rows. No cells are overwritten.|
|[PasteAsNestedTable](range-pasteasnestedtable-method-word.md)|Pastes a cell or group of cells as a nested table into the selected range.|
|[PasteExcelTable](range-pasteexceltable-method-word.md)|Pastes and formats a Microsoft Excel table.|
|[PasteSpecial](range-pastespecial-method-word.md)|Inserts the contents of the Clipboard. .|
|[PhoneticGuide](range-phoneticguide-method-word.md)|Adds phonetic guides to the specified range.|
|[Previous](range-previous-method-word.md)|Returns the previous range a relative to the specified range.|
|[PreviousSubdocument](range-previoussubdocument-method-word.md)|Moves the range to the previous subdocument.|
|[Relocate](range-relocate-method-word.md)|In outline view, moves the paragraphs within the specified range after the next visible paragraph or before the previous visible paragraph.|
|[Select](range-select-method-word.md)|Selects the specified range.|
|[SetListLevel](range-setlistlevel-method-word.md)|Sets the list level for one or more items in a numbered list.|
|[SetRange](range-setrange-method-word.md)|Sets the starting and ending character positions for an existing range.|
|[Sort](range-sort-method-word.md)|Sorts the paragraphs in the specified range.|
|[SortAscending](range-sortascending-method-word.md)|Sorts paragraphs or table rows in ascending alphanumeric order.|
|[SortByHeadings](range-sortbyheadings-method-word.md)|Sorts the headings in the specified range.|
|[SortDescending](range-sortdescending-method-word.md)|Sorts paragraphs in descending alphanumeric order.|
|[StartOf](range-startof-method-word.md)|Moves or extends the start position of the specified range or selection to the beginning of the nearest specified text unit. This method returns a  **Long** that indicates the number of characters by which the range or selection was moved or extended. The method returns a negative number if the movement is backward through the document.|
|[TCSCConverter](range-tcscconverter-method-word.md)|Converts the specified range from Traditional Chinese to Simplified Chinese or vice versa.|
|[WholeStory](range-wholestory-method-word.md)|Expands a range to include the entire story.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](range-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[Bold](range-bold-property-word.md)| **True** if the range is formatted as bold. Read/write **Long** .|
|[BoldBi](range-boldbi-property-word.md)| **True** if the font or range is formatted as bold. Returns **True** , **False** , or **wdUndefined** (for a mixture of bold and non-bold text). Can be set to **True** , **False** , or **wdToggle** . Read/write **Long** .|
|[BookmarkID](range-bookmarkid-property-word.md)|Returns the number of the bookmark that encloses the beginning of the specified range; returns 0 (zero) if there is no corresponding bookmark. Read-only  **Long** .|
|[Bookmarks](range-bookmarks-property-word.md)|Returns a  **[Bookmarks](bookmarks-object-word.md)** collection that represents all the bookmarks in a document, range, or selection. Read-only.|
|[Borders](range-borders-property-word.md)|Returns a  **[Borders](borders-object-word.md)** collection that represents all the borders for the specified object.|
|[Case](range-case-property-word.md)|Returns or sets a  **WdCharacterCase** constant that represents the case of the text in the specified range. Read/write.|
|[Cells](range-cells-property-word.md)|Returns a  **[Cells](cells-object-word.md)** collection that represents the table cells in a range. Read-only.|
|[Characters](range-characters-property-word.md)|Returns a  **[Characters](characters-object-word.md)** collection that represents the characters in a range. Read-only.|
|[CharacterStyle](range-characterstyle-property-word.md)|Returns a  **Variant** that represents the style used to format one or more characters. Read-only.|
|[CharacterWidth](range-characterwidth-property-word.md)|Returns or sets the character width of the specified range. Read/write  **WdCharacterWidth** .|
|[Columns](range-columns-property-word.md)|Returns a  **[Columns](columns-object-word.md)** collection that represents all the table columns in the range. Read-only.|
|[CombineCharacters](range-combinecharacters-property-word.md)| **True** if the specified range contains combined characters. Read/write **Boolean** .|
|[Comments](range-comments-property-word.md)|Returns a  **[Comments](comments-object-word.md)** collection that represents all the comments in the specified document, selection, or range. Read-only.|
|[Conflicts](range-conflicts-property-word.md)|Returns a [Conflicts](conflicts-object-word.md) collection object that contains all the conflict objects in the range. Read-only.|
|[ContentControls](range-contentcontrols-property-word.md)|Returns a  **[ContentControls](contentcontrols-object-word.md)** collection that represents the content controls contained within a range. Read-only.|
|[Creator](range-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[DisableCharacterSpaceGrid](range-disablecharacterspacegrid-property-word.md)| **True** if Microsoft Word ignores the number of characters per line for the corresponding **Range** object. Read/write **Boolean** .|
|[Document](range-document-property-word.md)|Returns a  **[Document](document-object-word.md)** object associated with the specified range. Read-only.|
|[Duplicate](range-duplicate-property-word.md)|Returns a read-only  **Range** object that represents all the properties of the specified range.|
|[Editors](range-editors-property-word.md)|Returns an  **Editors** object that represents all the users authorized to modify a selection or range within a document.|
|[EmphasisMark](range-emphasismark-property-word.md)|Returns or sets the emphasis mark for a character or designated character string. Read/write  **WdEmphasisMark** .|
|[End](range-end-property-word.md)|Returns or sets the ending character position of a range. Read/write  **Long** .|
|[EndnoteOptions](range-endnoteoptions-property-word.md)|Returns an  **EndnoteOptions** object that represents the endnotes in a range.|
|[Endnotes](range-endnotes-property-word.md)|Returns an  **[Endnotes](endnotes-object-word.md)** collection that represents all the endnotes in a range. Read-only.|
|[EnhMetaFileBits](range-enhmetafilebits-property-word.md)|Returns a  **Variant** that represents a picture representation of how a range of text appears.|
|[Fields](range-fields-property-word.md)|Returns a  **[Fields](fields-object-word.md)** collection that represents all the fields in the range. Read-only.|
|[Find](range-find-property-word.md)|Returns a  **[Find](find-object-word.md)** object that contains the criteria for a find operation. Read-only.|
|[FitTextWidth](range-fittextwidth-property-word.md)|Returns or sets the width (in the current measurement units) in which Microsoft Word fits the text in the current selection or range. Read/write  **Single** .|
|[Font](range-font-property-word.md)|Returns or sets a  **[Font](font-object-word.md)** object that represents the character formatting of the specified object. Read/write **Font** .|
|[FootnoteOptions](range-footnoteoptions-property-word.md)|Returns  **FootnoteOptions** object that represents the footnotes in a selection or range.|
|[Footnotes](range-footnotes-property-word.md)|Returns a  **[Footnotes](footnotes-object-word.md)** collection that represents all the footnotes in a range. Read-only.|
|[FormattedText](range-formattedtext-property-word.md)|Returns or sets a  **Range** object that includes the formatted text in the specified range or selection. Read/write.|
|[FormFields](range-formfields-property-word.md)|Returns a  **[FormFields](formfields-object-word.md)** collection that represents all the form fields in the range. Read-only.|
|[Frames](range-frames-property-word.md)|Returns a  **[Frames](frames-object-word.md)** collection that represents all the frames in a range. Read-only.|
|[GrammarChecked](range-grammarchecked-property-word.md)| **True** if a grammar check has been run on the specified range or document. Read/write **Boolean** .|
|[GrammaticalErrors](range-grammaticalerrors-property-word.md)|Returns a  **[ProofreadingErrors](proofreadingerrors-object-word.md)** collection that represents the sentences that failed the grammar check on the specified document or range. Read-only.|
|[HighlightColorIndex](range-highlightcolorindex-property-word.md)|Returns or sets the highlight color for the specified range. Read/write  **WdColorIndex** .|
|[HorizontalInVertical](range-horizontalinvertical-property-word.md)|Returns or sets the formatting for horizontal text set within vertical text. Read/write  **WdHorizontalInVerticalType** .|
|[HTMLDivisions](range-htmldivisions-property-word.md)|Returns an  **HTMLDivisions** object that represents an HTML division in a Web document.|
|[Hyperlinks](range-hyperlinks-property-word.md)|Returns a  **Hyperlinks** collection that represents all the hyperlinks in the specified range. Read-only.|
|[ID](range-id-property-word.md)|Returns or sets the identification name for the specified range. Read/write  **String** .|
|[Information](range-information-property-word.md)|Returns information about the specified range. Read-only  **Variant** .|
|[InlineShapes](range-inlineshapes-property-word.md)|Returns an  **InlineShapes** collection that represents all the **InlineShape** objects in a range. Read-only.|
|[IsEndOfRowMark](range-isendofrowmark-property-word.md)| **True** if the specified range is collapsed and is located at the end-of-row mark in a table. Read-only **Boolean** .|
|[Italic](range-italic-property-word.md)| **True** if the font or range is formatted as italic. Read/write **Long** .|
|[ItalicBi](range-italicbi-property-word.md)| **True** if the font or range is formatted as italic. Read/write **Long** .|
|[Kana](range-kana-property-word.md)|Returns or sets whether the specified range of Japanese language text is hiragana or katakana. Read/write  **WdKana** .|
|[LanguageDetected](range-languagedetected-property-word.md)|Returns or sets a value that specifies whether Microsoft Word has detected the language of the specified text. Read/write  **Boolean** .|
|[LanguageID](range-languageid-property-word.md)|Returns or sets a  **[WdLanguageID](wdlanguageid-enumeration-word.md)** constant that represents the language for the specified range. Read/write.|
|[LanguageIDFarEast](range-languageidfareast-property-word.md)|Returns or sets an East Asian language for the specified object. Read/write  **WdLanguageID** .|
|[LanguageIDOther](range-languageidother-property-word.md)|Returns or sets the language for the specified range. Read/write  **WdLanguageID** .|
|[ListFormat](range-listformat-property-word.md)|Returns a  **[ListFormat](listformat-object-word.md)** object that represents all the list formatting characteristics of a range. Read-only.|
|[ListParagraphs](range-listparagraphs-property-word.md)|Returns a  **ListParagraphs** collection that represents all the numbered paragraphs in the range. Read-only.|
|[ListStyle](range-liststyle-property-word.md)|Returns a  **Variant** that represents the style used to format a bulleted list or numbered list. Read-only.|
|[Locks](range-locks-property-word.md)|Returns a  **[CoAuthLocks](coauthlocks-object-word.md)** collection object that represents all the locks in the range. Read-only.|
|[NextStoryRange](range-nextstoryrange-property-word.md)|Returns a  **Range** object that refers to the next story. Read-only **Range** .|
|[NoProofing](range-noproofing-property-word.md)| **True** if the spelling and grammar checker ignores the specified text. Read/write **Long** .|
|[OMaths](range-omaths-property-word.md)|Returns an  **[OMaths](omaths-object-word.md)** collection that represents the **[OMath](omath-object-word.md)** objects within the specified range. Read-only.|
|[Orientation](range-orientation-property-word.md)|Returns or sets the orientation of text in a range when the Text Direction feature is enabled. Read/write  **WdTextOrientation** .|
|[PageSetup](range-pagesetup-property-word.md)|Returns a  **PageSetup** object that's associated with the specified range.|
|[ParagraphFormat](range-paragraphformat-property-word.md)|Returns or sets a  **[ParagraphFormat](paragraphformat-object-word.md)** object that represents the paragraph settings for the specified range. Read/write.|
|[Paragraphs](range-paragraphs-property-word.md)|Returns a  **Paragraphs** collection that represents all the paragraphs in the specified range. Read-only.|
|[ParagraphStyle](range-paragraphstyle-property-word.md)|Returns a  **Variant** that represents the style used to format a paragraph. Read-only.|
|[Parent](range-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **Range** object.|
|[ParentContentControl](range-parentcontentcontrol-property-word.md)|Returns a  **ContentControl** object that represents the parent content control for the specified range. Read-only.|
|[PreviousBookmarkID](range-previousbookmarkid-property-word.md)|Returns the number of the last bookmark that starts before or at the same place as the specified range. Read-only  **Long** .|
|[ReadabilityStatistics](range-readabilitystatistics-property-word.md)|Returns a  **ReadabilityStatistics** collection that represents the readability statistics for the specified document or range. Read-only.|
|[Revisions](range-revisions-property-word.md)|Returns a  **Revisions** collection that represents the tracked changes in the range. Read-only.|
|[Rows](range-rows-property-word.md)|Returns a  **Rows** collection that represents all the table rows in a range. Read-only.|
|[Scripts](range-scripts-property-word.md)|Returns a  **Scripts** collection that represents the collection of HTML scripts in the specified object.|
|[Sections](range-sections-property-word.md)|Returns a  **Sections** collection that represents the sections in the specified range. Read-only.|
|[Sentences](range-sentences-property-word.md)|Returns a  **Sentences** collection that represents all the sentences in the range. Read-only.|
|[Shading](range-shading-property-word.md)|Returns a  **Shading** object that refers to the shading formatting for the specified object.|
|[ShapeRange](range-shaperange-property-word.md)|Returns a  **[ShapeRange](shaperange-object-word.md)** collection that represents all the **Shape** objects in the specified range. Read-only.|
|[ShowAll](range-showall-property-word.md)| **True** if all nonprinting characters (such as hidden text, tab marks, space marks, and paragraph marks) are displayed. Read/write **Boolean** .|
|[SpellingChecked](range-spellingchecked-property-word.md)| **True** if spelling has been checked throughout the specified range or document. **False** if all or some of the range or document has not been checked for spelling. Read/write **Boolean** .|
|[SpellingErrors](range-spellingerrors-property-word.md)|Returns a  **ProofreadingErrors** collection that represents the words identified as spelling errors in the specified range. Read-only.|
|[Start](range-start-property-word.md)|Returns or sets the starting character position of a range. Read/write  **Long** .|
|[StoryLength](range-storylength-property-word.md)|Returns the number of characters in the story that contains the specified range. Read-only  **Long** .|
|[StoryType](range-storytype-property-word.md)|Returns the story type for the specified range, selection, or bookmark. Read-only  **[WdStoryType](wdstorytype-enumeration-word.md)** .|
|[Style](range-style-property-word.md)|Returns or sets the style for the specified object. Read/write  **Variant** .|
|[Subdocuments](range-subdocuments-property-word.md)|Returns a  **Subdocuments** collection that represents all the subdocuments in the specified range or document. Read-only.|
|[SynonymInfo](range-synonyminfo-property-word.md)|Returns a  **SynonymInfo** object that contains information from the thesaurus on synonyms, antonyms, or related words and expressions for the contents of a range.|
|[Tables](range-tables-property-word.md)|Returns a  **Tables** collection that represents all the tables in the specified range. Read-only.|
|[TableStyle](range-tablestyle-property-word.md)|Returns a  **Variant** that represents the style used to format a table. Read-only.|
|[Text](range-text-property-word.md)|Returns or sets the text in the specified range or selection. Read/write  **String** . Read/write **String** .|
|[TextRetrievalMode](range-textretrievalmode-property-word.md)|Returns a  **[TextRetrievalMode](textretrievalmode-object-word.md)** object that controls how text is retrieved from the specified **Range** . Read/write.|
|[TextVisibleOnScreen](range-textvisibleonscreen-property-word.md)|Returns a  **Long** that indicates whether the text in the specified range is visible on the screen. Read-only.|
|[TopLevelTables](range-topleveltables-property-word.md)|Returns a  **Tables** collection that represents the tables at the outermost nesting level in the current range. Read-only.|
|[TwoLinesInOne](range-twolinesinone-property-word.md)|Returns or sets whether Microsoft Word sets two lines of text in one and specifies the characters that enclose the text, if any. Read/write  **WdTwoLinesInOneType** .|
|[Underline](range-underline-property-word.md)|Returns or sets the type of underline applied to a range. Read/write  **WdUnderline** .|
|[Updates](range-updates-property-word.md)|Returns a [CoAuthUpdates](coauthupdates-object-word.md) collection object that represents all updates that were merged into the specified range at the last explicit save. Read-only.|
|[WordOpenXML](range-wordopenxml-property-word.md)|Returns a  **String** that represents the XML contained within the range in the Microsoft Word Open XML format. Read-only.|
|[Words](range-words-property-word.md)|Returns a  **Words** collection that represents all the words in a range. Read-only.|
|[XML](range-xml-property-word.md)|Returns a  **String** that represents the XML text in the specified object. .|

