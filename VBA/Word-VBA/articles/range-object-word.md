---
title: Range Object (Word)
keywords: vbawd10.chm2398
f1_keywords:
- vbawd10.chm2398
ms.prod: WORD
api_name:
- Word.Range
ms.assetid: 15a7a1c4-5f3f-5b6e-60e9-29688de3f274
---


# Range Object (Word)

Represents a contiguous area in a document. Each  **Range** object is defined by a starting and ending character position.


## Remarks

Similar to the way bookmarks are used in a document,  **Range** objects are used in Visual Basic procedures to identify specific portions of a document. However, unlike a bookmark, a **Range** object only exists while the procedure that defined it is running. **Range** objects are independent of the selection. That is, you can define and manipulate a range without changing the selection. You can also define multiple ranges in a document, while there can be only one selection per pane.

Use the  **Range** method to return a **Range** object defined by the given starting and ending character positions. The following example returns a **Range** object that refers to the first 10 characters in the active document.




```
Set myRange = ActiveDocument.Range(Start:=0, End:=10)
```

Use the  **Range** property to return a **Range** object defined by the beginning and end of another object. The **Range** property applies to many objects (for example, **Paragraph**, **Bookmark**, and **Cell** ). The following example returns a **Range** object that refers to the first paragraph in the active document.




```
Set aRange = ActiveDocument.Paragraphs(1).Range
```

The following example returns a  **Range** object that refers to the second through fourth paragraphs in the active document




```
Set aRange = ActiveDocument.Range( _ 
 Start:=ActiveDocument.Paragraphs(2).Range.Start, _ 
 End:=ActiveDocument.Paragraphs(4).Range.End)
```

For more information about working with  **Range** objects, see [Working with Range Objects](http://msdn.microsoft.com/library/working-with-range-objects%28Office.15%29.aspx).


## Methods



|**Name**|
|:-----|
|[AutoFormat](http://msdn.microsoft.com/library/range-autoformat-method-word%28Office.15%29.aspx)|
|[Calculate](http://msdn.microsoft.com/library/range-calculate-method-word%28Office.15%29.aspx)|
|[CheckGrammar](http://msdn.microsoft.com/library/range-checkgrammar-method-word%28Office.15%29.aspx)|
|[CheckSpelling](http://msdn.microsoft.com/library/range-checkspelling-method-word%28Office.15%29.aspx)|
|[CheckSynonyms](http://msdn.microsoft.com/library/range-checksynonyms-method-word%28Office.15%29.aspx)|
|[Collapse](http://msdn.microsoft.com/library/range-collapse-method-word%28Office.15%29.aspx)|
|[ComputeStatistics](http://msdn.microsoft.com/library/range-computestatistics-method-word%28Office.15%29.aspx)|
|[ConvertHangulAndHanja](http://msdn.microsoft.com/library/range-converthangulandhanja-method-word%28Office.15%29.aspx)|
|[ConvertToTable](http://msdn.microsoft.com/library/range-converttotable-method-word%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/range-copy-method-word%28Office.15%29.aspx)|
|[CopyAsPicture](http://msdn.microsoft.com/library/range-copyaspicture-method-word%28Office.15%29.aspx)|
|[Cut](http://msdn.microsoft.com/library/range-cut-method-word%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/range-delete-method-word%28Office.15%29.aspx)|
|[DetectLanguage](http://msdn.microsoft.com/library/range-detectlanguage-method-word%28Office.15%29.aspx)|
|[EndOf](http://msdn.microsoft.com/library/range-endof-method-word%28Office.15%29.aspx)|
|[Expand](http://msdn.microsoft.com/library/range-expand-method-word%28Office.15%29.aspx)|
|[ExportAsFixedFormat](http://msdn.microsoft.com/library/range-exportasfixedformat-method-word%28Office.15%29.aspx)|
|[ExportFragment](http://msdn.microsoft.com/library/range-exportfragment-method-word%28Office.15%29.aspx)|
|[GetSpellingSuggestions](http://msdn.microsoft.com/library/range-getspellingsuggestions-method-word%28Office.15%29.aspx)|
|[GoTo](http://msdn.microsoft.com/library/range-goto-method-word%28Office.15%29.aspx)|
|[GoToEditableRange](http://msdn.microsoft.com/library/range-gotoeditablerange-method-word%28Office.15%29.aspx)|
|[GoToNext](http://msdn.microsoft.com/library/range-gotonext-method-word%28Office.15%29.aspx)|
|[GoToPrevious](http://msdn.microsoft.com/library/range-gotoprevious-method-word%28Office.15%29.aspx)|
|[ImportFragment](http://msdn.microsoft.com/library/range-importfragment-method-word%28Office.15%29.aspx)|
|[InRange](http://msdn.microsoft.com/library/range-inrange-method-word%28Office.15%29.aspx)|
|[InsertAfter](http://msdn.microsoft.com/library/range-insertafter-method-word%28Office.15%29.aspx)|
|[InsertAlignmentTab](http://msdn.microsoft.com/library/range-insertalignmenttab-method-word%28Office.15%29.aspx)|
|[InsertAutoText](http://msdn.microsoft.com/library/range-insertautotext-method-word%28Office.15%29.aspx)|
|[InsertBefore](http://msdn.microsoft.com/library/range-insertbefore-method-word%28Office.15%29.aspx)|
|[InsertBreak](http://msdn.microsoft.com/library/range-insertbreak-method-word%28Office.15%29.aspx)|
|[InsertCaption](http://msdn.microsoft.com/library/range-insertcaption-method-word%28Office.15%29.aspx)|
|[InsertCrossReference](http://msdn.microsoft.com/library/range-insertcrossreference-method-word%28Office.15%29.aspx)|
|[InsertDatabase](http://msdn.microsoft.com/library/range-insertdatabase-method-word%28Office.15%29.aspx)|
|[InsertDateTime](http://msdn.microsoft.com/library/range-insertdatetime-method-word%28Office.15%29.aspx)|
|[InsertFile](http://msdn.microsoft.com/library/range-insertfile-method-word%28Office.15%29.aspx)|
|[InsertParagraph](http://msdn.microsoft.com/library/range-insertparagraph-method-word%28Office.15%29.aspx)|
|[InsertParagraphAfter](http://msdn.microsoft.com/library/range-insertparagraphafter-method-word%28Office.15%29.aspx)|
|[InsertParagraphBefore](http://msdn.microsoft.com/library/range-insertparagraphbefore-method-word%28Office.15%29.aspx)|
|[InsertSymbol](http://msdn.microsoft.com/library/range-insertsymbol-method-word%28Office.15%29.aspx)|
|[InsertXML](http://msdn.microsoft.com/library/range-insertxml-method-word%28Office.15%29.aspx)|
|[InStory](http://msdn.microsoft.com/library/range-instory-method-word%28Office.15%29.aspx)|
|[IsEqual](http://msdn.microsoft.com/library/range-isequal-method-word%28Office.15%29.aspx)|
|[LookupNameProperties](http://msdn.microsoft.com/library/range-lookupnameproperties-method-word%28Office.15%29.aspx)|
|[ModifyEnclosure](http://msdn.microsoft.com/library/range-modifyenclosure-method-word%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/range-move-method-word%28Office.15%29.aspx)|
|[MoveEnd](http://msdn.microsoft.com/library/range-moveend-method-word%28Office.15%29.aspx)|
|[MoveEndUntil](http://msdn.microsoft.com/library/range-moveenduntil-method-word%28Office.15%29.aspx)|
|[MoveEndWhile](http://msdn.microsoft.com/library/range-moveendwhile-method-word%28Office.15%29.aspx)|
|[MoveStart](http://msdn.microsoft.com/library/range-movestart-method-word%28Office.15%29.aspx)|
|[MoveStartUntil](http://msdn.microsoft.com/library/range-movestartuntil-method-word%28Office.15%29.aspx)|
|[MoveStartWhile](http://msdn.microsoft.com/library/range-movestartwhile-method-word%28Office.15%29.aspx)|
|[MoveUntil](http://msdn.microsoft.com/library/range-moveuntil-method-word%28Office.15%29.aspx)|
|[MoveWhile](http://msdn.microsoft.com/library/range-movewhile-method-word%28Office.15%29.aspx)|
|[Next](http://msdn.microsoft.com/library/range-next-method-word%28Office.15%29.aspx)|
|[NextSubdocument](http://msdn.microsoft.com/library/range-nextsubdocument-method-word%28Office.15%29.aspx)|
|[Paste](http://msdn.microsoft.com/library/range-paste-method-word%28Office.15%29.aspx)|
|[PasteAndFormat](http://msdn.microsoft.com/library/range-pasteandformat-method-word%28Office.15%29.aspx)|
|[PasteAppendTable](http://msdn.microsoft.com/library/range-pasteappendtable-method-word%28Office.15%29.aspx)|
|[PasteAsNestedTable](http://msdn.microsoft.com/library/range-pasteasnestedtable-method-word%28Office.15%29.aspx)|
|[PasteExcelTable](http://msdn.microsoft.com/library/range-pasteexceltable-method-word%28Office.15%29.aspx)|
|[PasteSpecial](http://msdn.microsoft.com/library/range-pastespecial-method-word%28Office.15%29.aspx)|
|[PhoneticGuide](http://msdn.microsoft.com/library/range-phoneticguide-method-word%28Office.15%29.aspx)|
|[Previous](http://msdn.microsoft.com/library/range-previous-method-word%28Office.15%29.aspx)|
|[PreviousSubdocument](http://msdn.microsoft.com/library/range-previoussubdocument-method-word%28Office.15%29.aspx)|
|[Relocate](http://msdn.microsoft.com/library/range-relocate-method-word%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/range-select-method-word%28Office.15%29.aspx)|
|[SetListLevel](http://msdn.microsoft.com/library/range-setlistlevel-method-word%28Office.15%29.aspx)|
|[SetRange](http://msdn.microsoft.com/library/range-setrange-method-word%28Office.15%29.aspx)|
|[Sort](http://msdn.microsoft.com/library/range-sort-method-word%28Office.15%29.aspx)|
|[SortAscending](http://msdn.microsoft.com/library/range-sortascending-method-word%28Office.15%29.aspx)|
|[SortByHeadings](http://msdn.microsoft.com/library/range-sortbyheadings-method-word%28Office.15%29.aspx)|
|[SortDescending](http://msdn.microsoft.com/library/range-sortdescending-method-word%28Office.15%29.aspx)|
|[StartOf](http://msdn.microsoft.com/library/range-startof-method-word%28Office.15%29.aspx)|
|[TCSCConverter](http://msdn.microsoft.com/library/range-tcscconverter-method-word%28Office.15%29.aspx)|
|[WholeStory](http://msdn.microsoft.com/library/range-wholestory-method-word%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/range-application-property-word%28Office.15%29.aspx)|
|[Bold](http://msdn.microsoft.com/library/range-bold-property-word%28Office.15%29.aspx)|
|[BoldBi](http://msdn.microsoft.com/library/range-boldbi-property-word%28Office.15%29.aspx)|
|[BookmarkID](http://msdn.microsoft.com/library/range-bookmarkid-property-word%28Office.15%29.aspx)|
|[Bookmarks](http://msdn.microsoft.com/library/range-bookmarks-property-word%28Office.15%29.aspx)|
|[Borders](http://msdn.microsoft.com/library/range-borders-property-word%28Office.15%29.aspx)|
|[Case](http://msdn.microsoft.com/library/range-case-property-word%28Office.15%29.aspx)|
|[Cells](http://msdn.microsoft.com/library/range-cells-property-word%28Office.15%29.aspx)|
|[Characters](http://msdn.microsoft.com/library/range-characters-property-word%28Office.15%29.aspx)|
|[CharacterStyle](http://msdn.microsoft.com/library/range-characterstyle-property-word%28Office.15%29.aspx)|
|[CharacterWidth](http://msdn.microsoft.com/library/range-characterwidth-property-word%28Office.15%29.aspx)|
|[Columns](http://msdn.microsoft.com/library/range-columns-property-word%28Office.15%29.aspx)|
|[CombineCharacters](http://msdn.microsoft.com/library/range-combinecharacters-property-word%28Office.15%29.aspx)|
|[Comments](http://msdn.microsoft.com/library/range-comments-property-word%28Office.15%29.aspx)|
|[Conflicts](http://msdn.microsoft.com/library/range-conflicts-property-word%28Office.15%29.aspx)|
|[ContentControls](http://msdn.microsoft.com/library/range-contentcontrols-property-word%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/range-creator-property-word%28Office.15%29.aspx)|
|[DisableCharacterSpaceGrid](http://msdn.microsoft.com/library/range-disablecharacterspacegrid-property-word%28Office.15%29.aspx)|
|[Document](http://msdn.microsoft.com/library/range-document-property-word%28Office.15%29.aspx)|
|[Duplicate](http://msdn.microsoft.com/library/range-duplicate-property-word%28Office.15%29.aspx)|
|[Editors](http://msdn.microsoft.com/library/range-editors-property-word%28Office.15%29.aspx)|
|[EmphasisMark](http://msdn.microsoft.com/library/range-emphasismark-property-word%28Office.15%29.aspx)|
|[End](http://msdn.microsoft.com/library/range-end-property-word%28Office.15%29.aspx)|
|[EndnoteOptions](http://msdn.microsoft.com/library/range-endnoteoptions-property-word%28Office.15%29.aspx)|
|[Endnotes](http://msdn.microsoft.com/library/range-endnotes-property-word%28Office.15%29.aspx)|
|[EnhMetaFileBits](http://msdn.microsoft.com/library/range-enhmetafilebits-property-word%28Office.15%29.aspx)|
|[Fields](http://msdn.microsoft.com/library/range-fields-property-word%28Office.15%29.aspx)|
|[Find](http://msdn.microsoft.com/library/range-find-property-word%28Office.15%29.aspx)|
|[FitTextWidth](http://msdn.microsoft.com/library/range-fittextwidth-property-word%28Office.15%29.aspx)|
|[Font](http://msdn.microsoft.com/library/range-font-property-word%28Office.15%29.aspx)|
|[FootnoteOptions](http://msdn.microsoft.com/library/range-footnoteoptions-property-word%28Office.15%29.aspx)|
|[Footnotes](http://msdn.microsoft.com/library/range-footnotes-property-word%28Office.15%29.aspx)|
|[FormattedText](http://msdn.microsoft.com/library/range-formattedtext-property-word%28Office.15%29.aspx)|
|[FormFields](http://msdn.microsoft.com/library/range-formfields-property-word%28Office.15%29.aspx)|
|[Frames](http://msdn.microsoft.com/library/range-frames-property-word%28Office.15%29.aspx)|
|[GrammarChecked](http://msdn.microsoft.com/library/range-grammarchecked-property-word%28Office.15%29.aspx)|
|[GrammaticalErrors](http://msdn.microsoft.com/library/range-grammaticalerrors-property-word%28Office.15%29.aspx)|
|[HighlightColorIndex](http://msdn.microsoft.com/library/range-highlightcolorindex-property-word%28Office.15%29.aspx)|
|[HorizontalInVertical](http://msdn.microsoft.com/library/range-horizontalinvertical-property-word%28Office.15%29.aspx)|
|[HTMLDivisions](http://msdn.microsoft.com/library/range-htmldivisions-property-word%28Office.15%29.aspx)|
|[Hyperlinks](http://msdn.microsoft.com/library/range-hyperlinks-property-word%28Office.15%29.aspx)|
|[ID](http://msdn.microsoft.com/library/range-id-property-word%28Office.15%29.aspx)|
|[Information](http://msdn.microsoft.com/library/range-information-property-word%28Office.15%29.aspx)|
|[InlineShapes](http://msdn.microsoft.com/library/range-inlineshapes-property-word%28Office.15%29.aspx)|
|[IsEndOfRowMark](http://msdn.microsoft.com/library/range-isendofrowmark-property-word%28Office.15%29.aspx)|
|[Italic](http://msdn.microsoft.com/library/range-italic-property-word%28Office.15%29.aspx)|
|[ItalicBi](http://msdn.microsoft.com/library/range-italicbi-property-word%28Office.15%29.aspx)|
|[Kana](http://msdn.microsoft.com/library/range-kana-property-word%28Office.15%29.aspx)|
|[LanguageDetected](http://msdn.microsoft.com/library/range-languagedetected-property-word%28Office.15%29.aspx)|
|[LanguageID](http://msdn.microsoft.com/library/range-languageid-property-word%28Office.15%29.aspx)|
|[LanguageIDFarEast](http://msdn.microsoft.com/library/range-languageidfareast-property-word%28Office.15%29.aspx)|
|[LanguageIDOther](http://msdn.microsoft.com/library/range-languageidother-property-word%28Office.15%29.aspx)|
|[ListFormat](http://msdn.microsoft.com/library/range-listformat-property-word%28Office.15%29.aspx)|
|[ListParagraphs](http://msdn.microsoft.com/library/range-listparagraphs-property-word%28Office.15%29.aspx)|
|[ListStyle](http://msdn.microsoft.com/library/range-liststyle-property-word%28Office.15%29.aspx)|
|[Locks](http://msdn.microsoft.com/library/range-locks-property-word%28Office.15%29.aspx)|
|[NextStoryRange](http://msdn.microsoft.com/library/range-nextstoryrange-property-word%28Office.15%29.aspx)|
|[NoProofing](http://msdn.microsoft.com/library/range-noproofing-property-word%28Office.15%29.aspx)|
|[OMaths](http://msdn.microsoft.com/library/range-omaths-property-word%28Office.15%29.aspx)|
|[Orientation](http://msdn.microsoft.com/library/range-orientation-property-word%28Office.15%29.aspx)|
|[PageSetup](http://msdn.microsoft.com/library/range-pagesetup-property-word%28Office.15%29.aspx)|
|[ParagraphFormat](http://msdn.microsoft.com/library/range-paragraphformat-property-word%28Office.15%29.aspx)|
|[Paragraphs](http://msdn.microsoft.com/library/range-paragraphs-property-word%28Office.15%29.aspx)|
|[ParagraphStyle](http://msdn.microsoft.com/library/range-paragraphstyle-property-word%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/range-parent-property-word%28Office.15%29.aspx)|
|[ParentContentControl](http://msdn.microsoft.com/library/range-parentcontentcontrol-property-word%28Office.15%29.aspx)|
|[PreviousBookmarkID](http://msdn.microsoft.com/library/range-previousbookmarkid-property-word%28Office.15%29.aspx)|
|[ReadabilityStatistics](http://msdn.microsoft.com/library/range-readabilitystatistics-property-word%28Office.15%29.aspx)|
|[Revisions](http://msdn.microsoft.com/library/range-revisions-property-word%28Office.15%29.aspx)|
|[Rows](http://msdn.microsoft.com/library/range-rows-property-word%28Office.15%29.aspx)|
|[Scripts](http://msdn.microsoft.com/library/range-scripts-property-word%28Office.15%29.aspx)|
|[Sections](http://msdn.microsoft.com/library/range-sections-property-word%28Office.15%29.aspx)|
|[Sentences](http://msdn.microsoft.com/library/range-sentences-property-word%28Office.15%29.aspx)|
|[Shading](http://msdn.microsoft.com/library/range-shading-property-word%28Office.15%29.aspx)|
|[ShapeRange](http://msdn.microsoft.com/library/range-shaperange-property-word%28Office.15%29.aspx)|
|[ShowAll](http://msdn.microsoft.com/library/range-showall-property-word%28Office.15%29.aspx)|
|[SpellingChecked](http://msdn.microsoft.com/library/range-spellingchecked-property-word%28Office.15%29.aspx)|
|[SpellingErrors](http://msdn.microsoft.com/library/range-spellingerrors-property-word%28Office.15%29.aspx)|
|[Start](http://msdn.microsoft.com/library/range-start-property-word%28Office.15%29.aspx)|
|[StoryLength](http://msdn.microsoft.com/library/range-storylength-property-word%28Office.15%29.aspx)|
|[StoryType](http://msdn.microsoft.com/library/range-storytype-property-word%28Office.15%29.aspx)|
|[Style](http://msdn.microsoft.com/library/range-style-property-word%28Office.15%29.aspx)|
|[Subdocuments](http://msdn.microsoft.com/library/range-subdocuments-property-word%28Office.15%29.aspx)|
|[SynonymInfo](http://msdn.microsoft.com/library/range-synonyminfo-property-word%28Office.15%29.aspx)|
|[Tables](http://msdn.microsoft.com/library/range-tables-property-word%28Office.15%29.aspx)|
|[TableStyle](http://msdn.microsoft.com/library/range-tablestyle-property-word%28Office.15%29.aspx)|
|[Text](http://msdn.microsoft.com/library/range-text-property-word%28Office.15%29.aspx)|
|[TextRetrievalMode](http://msdn.microsoft.com/library/range-textretrievalmode-property-word%28Office.15%29.aspx)|
|[TextVisibleOnScreen](http://msdn.microsoft.com/library/range-textvisibleonscreen-property-word%28Office.15%29.aspx)|
|[TopLevelTables](http://msdn.microsoft.com/library/range-topleveltables-property-word%28Office.15%29.aspx)|
|[TwoLinesInOne](http://msdn.microsoft.com/library/range-twolinesinone-property-word%28Office.15%29.aspx)|
|[Underline](http://msdn.microsoft.com/library/range-underline-property-word%28Office.15%29.aspx)|
|[Updates](http://msdn.microsoft.com/library/range-updates-property-word%28Office.15%29.aspx)|
|[WordOpenXML](http://msdn.microsoft.com/library/range-wordopenxml-property-word%28Office.15%29.aspx)|
|[Words](http://msdn.microsoft.com/library/range-words-property-word%28Office.15%29.aspx)|
|[XML](http://msdn.microsoft.com/library/range-xml-property-word%28Office.15%29.aspx)|

## See also


#### Other resources


<<<<<<< HEAD
=======

>>>>>>> d7667e83d23dbf8ebf5bf068ba6fed14c840c0f5
[Word Object Model Reference](http://msdn.microsoft.com/library/object-model-word-vba-reference%28Office.15%29.aspx)

