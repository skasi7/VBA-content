---
title: Selection Object (Word)
keywords: vbawd10.chm2421
f1_keywords:
- vbawd10.chm2421
ms.prod: WORD
api_name:
- Word.Selection
ms.assetid: 7b574a91-c33e-ecfd-6783-6b7528b2ed8f
---


# Selection Object (Word)

Represents the current selection in a window or pane. A selection represents either a selected (or highlighted) area in the document, or it represents the insertion point if nothing in the document is selected. There can be only one  **Selection** object per document window pane, and only one **Selection** object in the entire application can be active.


## Remarks

Use the  **Selection** property to return the **Selection** object. If no object qualifier is used with the **Selection** property, Microsoft Word returns the selection from the active pane of the active document window. The following example copies the current selection from the active document.


```
Selection.Copy
```

The following example deletes the selection from the third document in the  **Documents** collection. The document does not have to be active to access its current selection.




```
Documents(3).ActiveWindow.Selection.Cut
```

The following example copies the selection from the first pane of the active document and pastes it into the second pane.




```
ActiveDocument.ActiveWindow.Panes(1).Selection.Copy 
ActiveDocument.ActiveWindow.Panes(2).Selection.Paste
```

The  **Text** property is the default property of the **Selection** object. Use this property to set or return the text in the current selection. The following example assigns the text in the current selection to the variable `strTemp`, removing the last character if it is a paragraph mark.




```
Dim strTemp as String 
 
strTemp = Selection.Text 
If Right(strTemp, 1) = vbCr Then _ 
 strTemp = Left(strTemp, Len(strTemp) - 1)
```

The  **Selection** object has various methods and properties with which you can collapse, expand, or otherwise change the current selection. The following example moves the insertion point to the end of the document and selects the last three lines.




```
Selection.EndOf Unit:=wdStory, Extend:=wdMove 
Selection.HomeKey Unit:=wdLine, Extend:=wdExtend 
Selection.MoveUp Unit:=wdLine, Count:=2, Extend:=wdExtend
```

The  **Selection** object has various methods and properties with which you can edit selected text in a document. The following example selects the first sentence in the active document and replaces it with a new paragraph.




```
Options.ReplaceSelection = True 
ActiveDocument.Sentences(1).Select 
Selection.TypeText "Material below is confidential." 
Selection.TypeParagraph
```

The following example deletes the last paragraph of the first document in the  **Documents** collection and pastes it at the beginning of the second document.




```
With Documents(1) 
 .Paragraphs.Last.Range.Select 
 .ActiveWindow.Selection.Cut 
End With 
 
With Documents(2).ActiveWindow.Selection 
 .StartOf Unit:=wdStory, Extend:=wdMove 
 .Paste 
End With
```

The  **Selection** object has various methods and properties with which you can change the formatting of the current selection. The following example changes the font of the current selection from Times New Roman to Tahoma.




```
If Selection.Font.Name = "Times New Roman" Then _ 
 Selection.Font.Name = "Tahoma"
```

Use properties like  **Flags**, **Information**, and **Type** to return information about the current selection. You can use the following example in a procedure to determine whether there is anything selected in the active document; if there is not, the rest of the procedure is skipped.




```
If Selection.Type = wdSelectionIP Then 
 MsgBox Prompt:="You have not selected any text! Exiting procedure..." 
 Exit Sub 
End If
```

Even when a selection is collapsed to an insertion point, it is not necessarily empty. For example, the  **Text** property will still return the character to the right of the insertion point; this character also appears in the **Characters** collection of the **Selection** object. However, calling methods like **Cut** or **Copy** from a collapsed selection causes an error.

It is possible for the user to select a region in a document that does not represent contiguous text (for example, when using the ALT key with the mouse). Because the behavior of such a selection can be unpredictable, you may want to include a step in your code that checks the  **Type** property of a selection before performing any operations on it ( `Selection.Type = wdSelectionBlock`). Similarly, selections that include table cells can also lead to unpredictable behavior. The  **Information** property will tell you if a selection is inside a table ( `Selection.Information(wdWithinTable) = True`). The following example determines if a selection is normal (for example, it is not a row or column in a table, it is not a vertical block of text); you could use it to test the current selection before performing any operations on it.




```
If Selection.Type <> wdSelectionNormal Then 
 MsgBox Prompt:="Not a valid selection! Exiting procedure..." 
 Exit Sub 
End If
```

Because  **Range** objects share many of the same methods and properties as **Selection** objects, using **Range** objects is preferable for manipulating a document when there is not a reason to physically change the current selection. For more information about **Selection** and **Range** objects, see [Working with the Selection object](http://msdn.microsoft.com/library/working-with-the-selection-object%28Office.15%29.aspx) and [Working with Range objects](http://msdn.microsoft.com/library/working-with-range-objects%28Office.15%29.aspx).


## Methods



|**Name**|
|:-----|
|[BoldRun](http://msdn.microsoft.com/library/selection-boldrun-method-word%28Office.15%29.aspx)|
|[Calculate](http://msdn.microsoft.com/library/selection-calculate-method-word%28Office.15%29.aspx)|
|[ClearCharacterAllFormatting](http://msdn.microsoft.com/library/selection-clearcharacterallformatting-method-word%28Office.15%29.aspx)|
|[ClearCharacterDirectFormatting](http://msdn.microsoft.com/library/selection-clearcharacterdirectformatting-method-word%28Office.15%29.aspx)|
|[ClearCharacterStyle](http://msdn.microsoft.com/library/selection-clearcharacterstyle-method-word%28Office.15%29.aspx)|
|[ClearFormatting](http://msdn.microsoft.com/library/selection-clearformatting-method-word%28Office.15%29.aspx)|
|[ClearParagraphAllFormatting](http://msdn.microsoft.com/library/selection-clearparagraphallformatting-method-word%28Office.15%29.aspx)|
|[ClearParagraphDirectFormatting](http://msdn.microsoft.com/library/selection-clearparagraphdirectformatting-method-word%28Office.15%29.aspx)|
|[ClearParagraphStyle](http://msdn.microsoft.com/library/selection-clearparagraphstyle-method-word%28Office.15%29.aspx)|
|[Collapse](http://msdn.microsoft.com/library/selection-collapse-method-word%28Office.15%29.aspx)|
|[ConvertToTable](http://msdn.microsoft.com/library/selection-converttotable-method-word%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/selection-copy-method-word%28Office.15%29.aspx)|
|[CopyAsPicture](http://msdn.microsoft.com/library/selection-copyaspicture-method-word%28Office.15%29.aspx)|
|[CopyFormat](http://msdn.microsoft.com/library/selection-copyformat-method-word%28Office.15%29.aspx)|
|[CreateAutoTextEntry](http://msdn.microsoft.com/library/selection-createautotextentry-method-word%28Office.15%29.aspx)|
|[CreateTextbox](http://msdn.microsoft.com/library/selection-createtextbox-method-word%28Office.15%29.aspx)|
|[Cut](http://msdn.microsoft.com/library/selection-cut-method-word%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/selection-delete-method-word%28Office.15%29.aspx)|
|[DetectLanguage](http://msdn.microsoft.com/library/selection-detectlanguage-method-word%28Office.15%29.aspx)|
|[EndKey](http://msdn.microsoft.com/library/selection-endkey-method-word%28Office.15%29.aspx)|
|[EndOf](http://msdn.microsoft.com/library/selection-endof-method-word%28Office.15%29.aspx)|
|[EscapeKey](http://msdn.microsoft.com/library/selection-escapekey-method-word%28Office.15%29.aspx)|
|[Expand](http://msdn.microsoft.com/library/selection-expand-method-word%28Office.15%29.aspx)|
|[ExportAsFixedFormat](http://msdn.microsoft.com/library/selection-exportasfixedformat-method-word%28Office.15%29.aspx)|
|[Extend](http://msdn.microsoft.com/library/selection-extend-method-word%28Office.15%29.aspx)|
|[GoTo](http://msdn.microsoft.com/library/selection-goto-method-word%28Office.15%29.aspx)|
|[GoToEditableRange](http://msdn.microsoft.com/library/selection-gotoeditablerange-method-word%28Office.15%29.aspx)|
|[GoToNext](http://msdn.microsoft.com/library/selection-gotonext-method-word%28Office.15%29.aspx)|
|[GoToPrevious](http://msdn.microsoft.com/library/selection-gotoprevious-method-word%28Office.15%29.aspx)|
|[HomeKey](http://msdn.microsoft.com/library/selection-homekey-method-word%28Office.15%29.aspx)|
|[InRange](http://msdn.microsoft.com/library/selection-inrange-method-word%28Office.15%29.aspx)|
|[InsertAfter](http://msdn.microsoft.com/library/selection-insertafter-method-word%28Office.15%29.aspx)|
|[InsertBefore](http://msdn.microsoft.com/library/selection-insertbefore-method-word%28Office.15%29.aspx)|
|[InsertBreak](http://msdn.microsoft.com/library/selection-insertbreak-method-word%28Office.15%29.aspx)|
|[InsertCaption](http://msdn.microsoft.com/library/selection-insertcaption-method-word%28Office.15%29.aspx)|
|[InsertCells](http://msdn.microsoft.com/library/selection-insertcells-method-word%28Office.15%29.aspx)|
|[InsertColumns](http://msdn.microsoft.com/library/selection-insertcolumns-method-word%28Office.15%29.aspx)|
|[InsertColumnsRight](http://msdn.microsoft.com/library/selection-insertcolumnsright-method-word%28Office.15%29.aspx)|
|[InsertCrossReference](http://msdn.microsoft.com/library/selection-insertcrossreference-method-word%28Office.15%29.aspx)|
|[InsertDateTime](http://msdn.microsoft.com/library/selection-insertdatetime-method-word%28Office.15%29.aspx)|
|[InsertFile](http://msdn.microsoft.com/library/selection-insertfile-method-word%28Office.15%29.aspx)|
|[InsertFormula](http://msdn.microsoft.com/library/selection-insertformula-method-word%28Office.15%29.aspx)|
|[InsertNewPage](http://msdn.microsoft.com/library/selection-insertnewpage-method-word%28Office.15%29.aspx)|
|[InsertParagraph](http://msdn.microsoft.com/library/selection-insertparagraph-method-word%28Office.15%29.aspx)|
|[InsertParagraphAfter](http://msdn.microsoft.com/library/selection-insertparagraphafter-method-word%28Office.15%29.aspx)|
|[InsertParagraphBefore](http://msdn.microsoft.com/library/selection-insertparagraphbefore-method-word%28Office.15%29.aspx)|
|[InsertRows](http://msdn.microsoft.com/library/selection-insertrows-method-word%28Office.15%29.aspx)|
|[InsertRowsAbove](http://msdn.microsoft.com/library/selection-insertrowsabove-method-word%28Office.15%29.aspx)|
|[InsertRowsBelow](http://msdn.microsoft.com/library/selection-insertrowsbelow-method-word%28Office.15%29.aspx)|
|[InsertStyleSeparator](http://msdn.microsoft.com/library/selection-insertstyleseparator-method-word%28Office.15%29.aspx)|
|[InsertSymbol](http://msdn.microsoft.com/library/selection-insertsymbol-method-word%28Office.15%29.aspx)|
|[InsertXML](http://msdn.microsoft.com/library/selection-insertxml-method-word%28Office.15%29.aspx)|
|[InStory](http://msdn.microsoft.com/library/selection-instory-method-word%28Office.15%29.aspx)|
|[IsEqual](http://msdn.microsoft.com/library/selection-isequal-method-word%28Office.15%29.aspx)|
|[ItalicRun](http://msdn.microsoft.com/library/selection-italicrun-method-word%28Office.15%29.aspx)|
|[LtrPara](http://msdn.microsoft.com/library/selection-ltrpara-method-word%28Office.15%29.aspx)|
|[LtrRun](http://msdn.microsoft.com/library/selection-ltrrun-method-word%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/selection-move-method-word%28Office.15%29.aspx)|
|[MoveDown](http://msdn.microsoft.com/library/selection-movedown-method-word%28Office.15%29.aspx)|
|[MoveEnd](http://msdn.microsoft.com/library/selection-moveend-method-word%28Office.15%29.aspx)|
|[MoveEndUntil](http://msdn.microsoft.com/library/selection-moveenduntil-method-word%28Office.15%29.aspx)|
|[MoveEndWhile](http://msdn.microsoft.com/library/selection-moveendwhile-method-word%28Office.15%29.aspx)|
|[MoveLeft](http://msdn.microsoft.com/library/selection-moveleft-method-word%28Office.15%29.aspx)|
|[MoveRight](http://msdn.microsoft.com/library/selection-moveright-method-word%28Office.15%29.aspx)|
|[MoveStart](http://msdn.microsoft.com/library/selection-movestart-method-word%28Office.15%29.aspx)|
|[MoveStartUntil](http://msdn.microsoft.com/library/selection-movestartuntil-method-word%28Office.15%29.aspx)|
|[MoveStartWhile](http://msdn.microsoft.com/library/selection-movestartwhile-method-word%28Office.15%29.aspx)|
|[MoveUntil](http://msdn.microsoft.com/library/selection-moveuntil-method-word%28Office.15%29.aspx)|
|[MoveUp](http://msdn.microsoft.com/library/selection-moveup-method-word%28Office.15%29.aspx)|
|[MoveWhile](http://msdn.microsoft.com/library/selection-movewhile-method-word%28Office.15%29.aspx)|
|[Next](http://msdn.microsoft.com/library/selection-next-method-word%28Office.15%29.aspx)|
|[NextField](http://msdn.microsoft.com/library/selection-nextfield-method-word%28Office.15%29.aspx)|
|[NextRevision](http://msdn.microsoft.com/library/selection-nextrevision-method-word%28Office.15%29.aspx)|
|[NextSubdocument](http://msdn.microsoft.com/library/selection-nextsubdocument-method-word%28Office.15%29.aspx)|
|[Paste](http://msdn.microsoft.com/library/selection-paste-method-word%28Office.15%29.aspx)|
|[PasteAndFormat](http://msdn.microsoft.com/library/selection-pasteandformat-method-word%28Office.15%29.aspx)|
|[PasteAppendTable](http://msdn.microsoft.com/library/selection-pasteappendtable-method-word%28Office.15%29.aspx)|
|[PasteAsNestedTable](http://msdn.microsoft.com/library/selection-pasteasnestedtable-method-word%28Office.15%29.aspx)|
|[PasteExcelTable](http://msdn.microsoft.com/library/selection-pasteexceltable-method-word%28Office.15%29.aspx)|
|[PasteFormat](http://msdn.microsoft.com/library/selection-pasteformat-method-word%28Office.15%29.aspx)|
|[PasteSpecial](http://msdn.microsoft.com/library/selection-pastespecial-method-word%28Office.15%29.aspx)|
|[Previous](http://msdn.microsoft.com/library/selection-previous-method-word%28Office.15%29.aspx)|
|[PreviousField](http://msdn.microsoft.com/library/selection-previousfield-method-word%28Office.15%29.aspx)|
|[PreviousRevision](http://msdn.microsoft.com/library/selection-previousrevision-method-word%28Office.15%29.aspx)|
|[PreviousSubdocument](http://msdn.microsoft.com/library/selection-previoussubdocument-method-word%28Office.15%29.aspx)|
|[ReadingModeGrowFont](http://msdn.microsoft.com/library/selection-readingmodegrowfont-method-word%28Office.15%29.aspx)|
|[ReadingModeShrinkFont](http://msdn.microsoft.com/library/selection-readingmodeshrinkfont-method-word%28Office.15%29.aspx)|
|[RtlPara](http://msdn.microsoft.com/library/selection-rtlpara-method-word%28Office.15%29.aspx)|
|[RtlRun](http://msdn.microsoft.com/library/selection-rtlrun-method-word%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/selection-select-method-word%28Office.15%29.aspx)|
|[SelectCell](http://msdn.microsoft.com/library/selection-selectcell-method-word%28Office.15%29.aspx)|
|[SelectColumn](http://msdn.microsoft.com/library/selection-selectcolumn-method-word%28Office.15%29.aspx)|
|[SelectCurrentAlignment](http://msdn.microsoft.com/library/selection-selectcurrentalignment-method-word%28Office.15%29.aspx)|
|[SelectCurrentColor](http://msdn.microsoft.com/library/selection-selectcurrentcolor-method-word%28Office.15%29.aspx)|
|[SelectCurrentFont](http://msdn.microsoft.com/library/selection-selectcurrentfont-method-word%28Office.15%29.aspx)|
|[SelectCurrentIndent](http://msdn.microsoft.com/library/selection-selectcurrentindent-method-word%28Office.15%29.aspx)|
|[SelectCurrentSpacing](http://msdn.microsoft.com/library/selection-selectcurrentspacing-method-word%28Office.15%29.aspx)|
|[SelectCurrentTabs](http://msdn.microsoft.com/library/selection-selectcurrenttabs-method-word%28Office.15%29.aspx)|
|[SelectRow](http://msdn.microsoft.com/library/selection-selectrow-method-word%28Office.15%29.aspx)|
|[SetRange](http://msdn.microsoft.com/library/selection-setrange-method-word%28Office.15%29.aspx)|
|[Shrink](http://msdn.microsoft.com/library/selection-shrink-method-word%28Office.15%29.aspx)|
|[ShrinkDiscontiguousSelection](http://msdn.microsoft.com/library/selection-shrinkdiscontiguousselection-method-word%28Office.15%29.aspx)|
|[Sort](http://msdn.microsoft.com/library/selection-sort-method-word%28Office.15%29.aspx)|
|[SortAscending](http://msdn.microsoft.com/library/selection-sortascending-method-word%28Office.15%29.aspx)|
|[SortByHeadings](http://msdn.microsoft.com/library/selection-sortbyheadings-method-word%28Office.15%29.aspx)|
|[SortDescending](http://msdn.microsoft.com/library/selection-sortdescending-method-word%28Office.15%29.aspx)|
|[SplitTable](http://msdn.microsoft.com/library/selection-splittable-method-word%28Office.15%29.aspx)|
|[StartOf](http://msdn.microsoft.com/library/selection-startof-method-word%28Office.15%29.aspx)|
|[ToggleCharacterCode](http://msdn.microsoft.com/library/selection-togglecharactercode-method-word%28Office.15%29.aspx)|
|[TypeBackspace](http://msdn.microsoft.com/library/selection-typebackspace-method-word%28Office.15%29.aspx)|
|[TypeParagraph](http://msdn.microsoft.com/library/selection-typeparagraph-method-word%28Office.15%29.aspx)|
|[TypeText](http://msdn.microsoft.com/library/selection-typetext-method-word%28Office.15%29.aspx)|
|[WholeStory](http://msdn.microsoft.com/library/selection-wholestory-method-word%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Active](http://msdn.microsoft.com/library/selection-active-property-word%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/selection-application-property-word%28Office.15%29.aspx)|
|[BookmarkID](http://msdn.microsoft.com/library/selection-bookmarkid-property-word%28Office.15%29.aspx)|
|[Bookmarks](http://msdn.microsoft.com/library/selection-bookmarks-property-word%28Office.15%29.aspx)|
|[Borders](http://msdn.microsoft.com/library/selection-borders-property-word%28Office.15%29.aspx)|
|[Cells](http://msdn.microsoft.com/library/selection-cells-property-word%28Office.15%29.aspx)|
|[Characters](http://msdn.microsoft.com/library/selection-characters-property-word%28Office.15%29.aspx)|
|[ChildShapeRange](http://msdn.microsoft.com/library/selection-childshaperange-property-word%28Office.15%29.aspx)|
|[Columns](http://msdn.microsoft.com/library/selection-columns-property-word%28Office.15%29.aspx)|
|[ColumnSelectMode](http://msdn.microsoft.com/library/selection-columnselectmode-property-word%28Office.15%29.aspx)|
|[Comments](http://msdn.microsoft.com/library/selection-comments-property-word%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/selection-creator-property-word%28Office.15%29.aspx)|
|[Document](http://msdn.microsoft.com/library/selection-document-property-word%28Office.15%29.aspx)|
|[Editors](http://msdn.microsoft.com/library/selection-editors-property-word%28Office.15%29.aspx)|
|[End](http://msdn.microsoft.com/library/selection-end-property-word%28Office.15%29.aspx)|
|[EndnoteOptions](http://msdn.microsoft.com/library/selection-endnoteoptions-property-word%28Office.15%29.aspx)|
|[Endnotes](http://msdn.microsoft.com/library/selection-endnotes-property-word%28Office.15%29.aspx)|
|[EnhMetaFileBits](http://msdn.microsoft.com/library/selection-enhmetafilebits-property-word%28Office.15%29.aspx)|
|[ExtendMode](http://msdn.microsoft.com/library/selection-extendmode-property-word%28Office.15%29.aspx)|
|[Fields](http://msdn.microsoft.com/library/selection-fields-property-word%28Office.15%29.aspx)|
|[Find](http://msdn.microsoft.com/library/selection-find-property-word%28Office.15%29.aspx)|
|[FitTextWidth](http://msdn.microsoft.com/library/selection-fittextwidth-property-word%28Office.15%29.aspx)|
|[Flags](http://msdn.microsoft.com/library/selection-flags-property-word%28Office.15%29.aspx)|
|[Font](http://msdn.microsoft.com/library/selection-font-property-word%28Office.15%29.aspx)|
|[FootnoteOptions](http://msdn.microsoft.com/library/selection-footnoteoptions-property-word%28Office.15%29.aspx)|
|[Footnotes](http://msdn.microsoft.com/library/selection-footnotes-property-word%28Office.15%29.aspx)|
|[FormattedText](http://msdn.microsoft.com/library/selection-formattedtext-property-word%28Office.15%29.aspx)|
|[FormFields](http://msdn.microsoft.com/library/selection-formfields-property-word%28Office.15%29.aspx)|
|[Frames](http://msdn.microsoft.com/library/selection-frames-property-word%28Office.15%29.aspx)|
|[HasChildShapeRange](http://msdn.microsoft.com/library/selection-haschildshaperange-property-word%28Office.15%29.aspx)|
|[HeaderFooter](http://msdn.microsoft.com/library/selection-headerfooter-property-word%28Office.15%29.aspx)|
|[HTMLDivisions](http://msdn.microsoft.com/library/selection-htmldivisions-property-word%28Office.15%29.aspx)|
|[Hyperlinks](http://msdn.microsoft.com/library/selection-hyperlinks-property-word%28Office.15%29.aspx)|
|[Information](http://msdn.microsoft.com/library/selection-information-property-word%28Office.15%29.aspx)|
|[InlineShapes](http://msdn.microsoft.com/library/selection-inlineshapes-property-word%28Office.15%29.aspx)|
|[IPAtEndOfLine](http://msdn.microsoft.com/library/selection-ipatendofline-property-word%28Office.15%29.aspx)|
|[IsEndOfRowMark](http://msdn.microsoft.com/library/selection-isendofrowmark-property-word%28Office.15%29.aspx)|
|[LanguageDetected](http://msdn.microsoft.com/library/selection-languagedetected-property-word%28Office.15%29.aspx)|
|[LanguageID](http://msdn.microsoft.com/library/selection-languageid-property-word%28Office.15%29.aspx)|
|[LanguageIDFarEast](http://msdn.microsoft.com/library/selection-languageidfareast-property-word%28Office.15%29.aspx)|
|[LanguageIDOther](http://msdn.microsoft.com/library/selection-languageidother-property-word%28Office.15%29.aspx)|
|[NoProofing](http://msdn.microsoft.com/library/selection-noproofing-property-word%28Office.15%29.aspx)|
|[OMaths](http://msdn.microsoft.com/library/selection-omaths-property-word%28Office.15%29.aspx)|
|[Orientation](http://msdn.microsoft.com/library/selection-orientation-property-word%28Office.15%29.aspx)|
|[PageSetup](http://msdn.microsoft.com/library/selection-pagesetup-property-word%28Office.15%29.aspx)|
|[ParagraphFormat](http://msdn.microsoft.com/library/selection-paragraphformat-property-word%28Office.15%29.aspx)|
|[Paragraphs](http://msdn.microsoft.com/library/selection-paragraphs-property-word%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/selection-parent-property-word%28Office.15%29.aspx)|
|[PreviousBookmarkID](http://msdn.microsoft.com/library/selection-previousbookmarkid-property-word%28Office.15%29.aspx)|
|[Range](http://msdn.microsoft.com/library/selection-range-property-word%28Office.15%29.aspx)|
|[Rows](http://msdn.microsoft.com/library/selection-rows-property-word%28Office.15%29.aspx)|
|[Sections](http://msdn.microsoft.com/library/selection-sections-property-word%28Office.15%29.aspx)|
|[Sentences](http://msdn.microsoft.com/library/selection-sentences-property-word%28Office.15%29.aspx)|
|[Shading](http://msdn.microsoft.com/library/selection-shading-property-word%28Office.15%29.aspx)|
|[ShapeRange](http://msdn.microsoft.com/library/selection-shaperange-property-word%28Office.15%29.aspx)|
|[Start](http://msdn.microsoft.com/library/selection-start-property-word%28Office.15%29.aspx)|
|[StartIsActive](http://msdn.microsoft.com/library/selection-startisactive-property-word%28Office.15%29.aspx)|
|[StoryLength](http://msdn.microsoft.com/library/selection-storylength-property-word%28Office.15%29.aspx)|
|[StoryType](http://msdn.microsoft.com/library/selection-storytype-property-word%28Office.15%29.aspx)|
|[Style](http://msdn.microsoft.com/library/selection-style-property-word%28Office.15%29.aspx)|
|[Tables](http://msdn.microsoft.com/library/selection-tables-property-word%28Office.15%29.aspx)|
|[Text](http://msdn.microsoft.com/library/selection-text-property-word%28Office.15%29.aspx)|
|[TopLevelTables](http://msdn.microsoft.com/library/selection-topleveltables-property-word%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/selection-type-property-word%28Office.15%29.aspx)|
|[WordOpenXML](http://msdn.microsoft.com/library/selection-wordopenxml-property-word%28Office.15%29.aspx)|
|[Words](http://msdn.microsoft.com/library/selection-words-property-word%28Office.15%29.aspx)|
|[XML](http://msdn.microsoft.com/library/selection-xml-property-word%28Office.15%29.aspx)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/object-model-word-vba-reference%28Office.15%29.aspx)
<<<<<<< HEAD
=======

>>>>>>> d7667e83d23dbf8ebf5bf068ba6fed14c840c0f5

