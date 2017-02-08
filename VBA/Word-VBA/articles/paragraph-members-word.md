---
title: Paragraph Members (Word)
ms.prod: WORD
ms.assetid: e1fc5b91-e908-580e-ab72-898648a5c0c3
---


# Paragraph Members (Word)
Represents a single paragraph in a selection, range, or document. The  **Paragraph** object is a member of the **[Paragraphs](paragraphs-object-word.md)** collection. The **Paragraphs** collection includes all the paragraphs in a selection, range, or document.

Represents a single paragraph in a selection, range, or document. The  **Paragraph** object is a member of the **[Paragraphs](paragraphs-object-word.md)** collection. The **Paragraphs** collection includes all the paragraphs in a selection, range, or document.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[CloseUp](paragraph-closeup-method-word.md)|Removes any spacing before the specified paragraph.|
|[Indent](paragraph-indent-method-word.md)|Indents one or more paragraphs by one level.|
|[IndentCharWidth](paragraph-indentcharwidth-method-word.md)|Indents a paragraphs by a specified number of characters.|
|[IndentFirstLineCharWidth](paragraph-indentfirstlinecharwidth-method-word.md)|Indents the first line of one or more paragraphs by a specified number of characters.|
|[JoinList](paragraph-joinlist-method-word.md)|Joins a list paragraph with the closest list above or below the specified paragraph.|
|[ListAdvanceTo](paragraph-listadvanceto-method-word.md)|Sets the list levels for a paragraph in a list.|
|[Next](paragraph-next-method-word.md)|Returns a  **Paragraph** object that represents the next paragraph.|
|[OpenOrCloseUp](paragraph-openorcloseup-method-word.md)|Toggles the spacing before a paragraph.|
|[OpenUp](paragraph-openup-method-word.md)|Sets spacing before the specified paragraphs to 12 points.|
|[Outdent](paragraph-outdent-method-word.md)|Removes one level of indent for one or more paragraphs.|
|[OutlineDemote](paragraph-outlinedemote-method-word.md)|Applies the next heading level style (Heading 1 through Heading 8) to the specified paragraph or paragraphs.|
|[OutlineDemoteToBody](paragraph-outlinedemotetobody-method-word.md)|Demotes the specified paragraph to body text by applying the Normal style.|
|[OutlinePromote](paragraph-outlinepromote-method-word.md)|Applies the previous heading level style (Heading 1 through Heading 8) to the specified paragraph or paragraphs.|
|[Previous](paragraph-previous-method-word.md)|Returns the previous paragraph as a  **Paragraph** object.|
|[Reset](paragraph-reset-method-word.md)|Removes manual paragraph formatting (formatting not applied using a style).|
|[ResetAdvanceTo](paragraph-resetadvanceto-method-word.md)|Resets a paragraph that uses custom list levels to the original level settings.|
|[SelectNumber](paragraph-selectnumber-method-word.md)|Selects the number or bullet in a list.|
|[SeparateList](paragraph-separatelist-method-word.md)|Separates a list into two separate lists. For numbered lists, the new list restarts numbering at the starting number, usually 1.|
|[Space1](paragraph-space1-method-word.md)|Single-spaces the specified paragraphs.|
|[Space15](paragraph-space15-method-word.md)|Formats the specified paragraphs with 1.5-line spacing.|
|[Space2](paragraph-space2-method-word.md)|Double-spaces the specified paragraphs.|
|[TabHangingIndent](paragraph-tabhangingindent-method-word.md)|Sets a hanging indent to a specified number of tab stops. .|
|[TabIndent](paragraph-tabindent-method-word.md)|Sets the left indent for the specified paragraphs to a specified number of tab stops. .|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[AddSpaceBetweenFarEastAndAlpha](paragraph-addspacebetweenfareastandalpha-property-word.md)| **True** if Microsoft Word is set to automatically add spaces between Japanese and Latin text for the specified paragraphs. This property returns **wdUndefined** if it's set to **True** for only some of the specified paragraphs. Read/write **Long** .|
|[AddSpaceBetweenFarEastAndDigit](paragraph-addspacebetweenfareastanddigit-property-word.md)| **True** if Microsoft Word is set to automatically add spaces between Japanese text and numbers for the specified paragraphs. This property returns **wdUndefined** if it's set to **True** for only some of the specified paragraphs. Read/write **Long** .|
|[Alignment](paragraph-alignment-property-word.md)|Returns or sets a  **WdParagraphAlignment** constant that represents the alignment for the specified paragraphs. Read/write.|
|[Application](paragraph-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[AutoAdjustRightIndent](paragraph-autoadjustrightindent-property-word.md)| **True** if Microsoft Word is set to automatically adjust the right indent for the specified paragraphs if you've specified a set number of characters per line. Returns **wdUndefined** if the **AutoAdjustRightIndent** property is set to **True** for only some of the specified paragraphs. Read/write **Long** .|
|[BaseLineAlignment](paragraph-baselinealignment-property-word.md)|Returns or sets a  **WdBaselineAlignment** constant that represents the vertical position of fonts on a line. Read/write.|
|[Borders](paragraph-borders-property-word.md)|Returns a  **[Borders](borders-object-word.md)** collection that represents all the borders for the specified paragraph.|
|[CharacterUnitFirstLineIndent](paragraph-characterunitfirstlineindent-property-word.md)|Returns or sets the value (in characters) for a first-line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent. Read/write  **Single** .|
|[CharacterUnitLeftIndent](paragraph-characterunitleftindent-property-word.md)|Returns or sets the left indent value (in characters) for the specified paragraphs. Read/write  **Single** .|
|[CharacterUnitRightIndent](paragraph-characterunitrightindent-property-word.md)|Returns or sets the right indent value (in characters) for the specified paragraphs. Read/write  **Single** .|
|[CollapsedState](paragraph-collapsedstate-property-word.md)|Returns or sets whether the specified paragraph is currently in a collapsed state. Read-write  **Boolean**.|
|[CollapseHeadingByDefault](paragraph-collapseheadingbydefault-property-word.md)|Returns or sets whether the specified paragraph is collapsed by default when the document loads. Read-write  **Boolean**.|
|[Creator](paragraph-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[DisableLineHeightGrid](paragraph-disablelineheightgrid-property-word.md)| **True** if Microsoft Word aligns characters in the specified paragraphs to the line grid when a set number of lines per page is specified. Returns **wdUndefined** if the **DisableLineHeightGrid** property is set to **True** for only some of the specified paragraphs. Read/write **Long** .|
|[DropCap](paragraph-dropcap-property-word.md)|Returns a  **[DropCap](dropcap-object-word.md)** object that represents a dropped capital letter for the specified paragraph. Read-only.|
|[FarEastLineBreakControl](paragraph-fareastlinebreakcontrol-property-word.md)| **True** if Microsoft Word applies East Asian line-breaking rules to the specified paragraphs. Returns **wdUndefined** if the **FarEastLineBreakControl** property is set to **True** for only some of the specified paragraphs. Read/write **Long** .|
|[FirstLineIndent](paragraph-firstlineindent-property-word.md)|Returns or sets the value (in points) for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent. Read/write  **Single** .|
|[Format](paragraph-format-property-word.md)|Returns or sets a  **[ParagraphFormat](paragraphformat-object-word.md)** object that represents the formatting of the specified paragraph or paragraphs.|
|[HalfWidthPunctuationOnTopOfLine](paragraph-halfwidthpunctuationontopofline-property-word.md)| **True** if Microsoft Word changes punctuation symbols at the beginning of a line to half-width characters for the specified paragraphs. This property returns **wdUndefined** if it's set to **True** for only some of the specified paragraphs. Read/write **Long** .|
|[HangingPunctuation](paragraph-hangingpunctuation-property-word.md)| **True** if hanging punctuation is enabled for the specified paragraphs. This property returns **wdUndefined** if it's set to **True** for only some of the specified paragraphs. Read/write **Long** .|
|[Hyphenation](paragraph-hyphenation-property-word.md)| **True** if the specified paragraphs are included in automatic hyphenation. **False** if the specified paragraphs are to be excluded from automatic hyphenation. Read/write **Long** .|
|[ID](paragraph-id-property-word.md)|Returns or sets the identifying label for the specified object when the current document is saved as a Web page. Read/write  **String** .|
|[IsStyleSeparator](paragraph-isstyleseparator-property-word.md)| **True** if a paragraph contains a special hidden paragraph mark that allows Microsoft Word to appear to join paragraphs of different paragraph styles. Read-only **Boolean** .|
|[KeepTogether](paragraph-keeptogether-property-word.md)| **True** if all lines in the specified paragraph remain on the same page when Microsoft Word repaginates the document. Read/write **Long** .|
|[KeepWithNext](paragraph-keepwithnext-property-word.md)| **True** if the specified paragraph remains on the same page as the paragraph that follows it when Microsoft Word repaginates the document. Read/write **Long** .|
|[LeftIndent](paragraph-leftindent-property-word.md)|Returns or sets a  **Single** that represents the left indent value (in points) for the specified paragraph. Read/write.|
|[LineSpacing](paragraph-linespacing-property-word.md)|Returns or sets the line spacing (in points) for the specified paragraphs. Read/write  **Single** .|
|[LineSpacingRule](paragraph-linespacingrule-property-word.md)|Returns or sets the line spacing for the specified paragraph. Read/write  **[WdLineSpacing](wdlinespacing-enumeration-word.md)** .|
|[LineUnitAfter](paragraph-lineunitafter-property-word.md)|Returns or sets the amount of spacing (in gridlines) after the specified paragraph. Read/write  **Single** .|
|[LineUnitBefore](paragraph-lineunitbefore-property-word.md)|Returns or sets the amount of spacing (in gridlines) before the specified paragraph. Read/write  **Single** .|
|[ListNumberOriginal](paragraph-listnumberoriginal-property-word.md)|Returns an  **Integer** that represents the original list level for a paragraph. Read-only.|
|[MirrorIndents](paragraph-mirrorindents-property-word.md)|Returns or sets a  **Long** that represents whether left and right indents are the same width. Can be **True** , **False** , or **wdUndefined** . Read/write.|
|[NoLineNumber](paragraph-nolinenumber-property-word.md)| **True** if line numbers are repressed for the specified paragraph. Read/write **Long** .|
|[OutlineLevel](paragraph-outlinelevel-property-word.md)|Returns or sets the outline level for the specified paragraph. Read/write  **[WdOutlineLevel](wdoutlinelevel-enumeration-word.md)** .|
|[PageBreakBefore](paragraph-pagebreakbefore-property-word.md)| **True** if a page break is forced before the specified paragraphs. Read/write **Long** .|
|[Parent](paragraph-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **Paragraph** object.|
|[Range](paragraph-range-property-word.md)|Returns a  **Range** object that represents the portion of a document that is contained within the specified paragraph.|
|[ReadingOrder](paragraph-readingorder-property-word.md)|Returns or sets the reading order of the specified paragraph without changing the alignment. Read/write  **WdReadingOrder** .|
|[RightIndent](paragraph-rightindent-property-word.md)|Returns or sets the right indent (in points) for the specified paragraph. Read/write  **Single** .|
|[Shading](paragraph-shading-property-word.md)|Returns a  **[Shading](shading-object-word.md)** object that refers to the shading formatting for the specified paragraph.|
|[SpaceAfter](paragraph-spaceafter-property-word.md)|Returns or sets the amount of spacing (in points) after the specified paragraph or text column. Read/write  **Single** .|
|[SpaceAfterAuto](paragraph-spaceafterauto-property-word.md)| **True** if Microsoft Word automatically sets the amount of spacing after the specified paragraphs. Read/write **Long** .|
|[SpaceBefore](paragraph-spacebefore-property-word.md)|Returns or sets the spacing (in points) before the specified paragraphs. Read/write  **Single** .|
|[SpaceBeforeAuto](paragraph-spacebeforeauto-property-word.md)| **True** if Microsoft Word automatically sets the amount of spacing before the specified paragraphs. Read/write **Long** .|
|[Style](paragraph-style-property-word.md)|Returns or sets the style for the specified object. Read/write  **Variant** .|
|[TabStops](paragraph-tabstops-property-word.md)|Returns or sets a  **TabStops** collection that represents all the custom tab stops for the specified paragraph. Read/write.|
|[TextboxTightWrap](paragraph-textboxtightwrap-property-word.md)|Returns or sets a  **[WdTextboxTightWrap](wdtextboxtightwrap-enumeration-word.md)** constant that represents how tightly text wraps around shapes or text boxes. Read/write.|
|[WidowControl](paragraph-widowcontrol-property-word.md)| **True** if the first and last lines in the specified paragraph remain on the same page as the rest of the paragraph when Word repaginates the document. Read/write **Long** .|
|[WordWrap](paragraph-wordwrap-property-word.md)| **True** if Microsoft Word wraps Latin text in the middle of a word in the specified paragraphs or text frames. Read/write **Long** .|

