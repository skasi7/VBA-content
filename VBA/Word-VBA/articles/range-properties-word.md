---
title: Range Properties (Word)
ms.prod: WORD
ms.assetid: 4639f5ff-36fa-493f-9ce0-54862d5bb5c8
---


# Range Properties (Word)

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

