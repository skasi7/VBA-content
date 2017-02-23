---
title: ParagraphFormat Members (Word)
ms.prod: WORD
ms.assetid: d34122e7-adfb-dd34-eb1d-cd62b20a83ff
---


# ParagraphFormat Members (Word)
Represents all the formatting for a paragraph.

Represents all the formatting for a paragraph.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[CloseUp](paragraphformat-closeup-method-word.md)|Removes any spacing before paragraphs in the specified paragraph format.|
|[IndentCharWidth](paragraphformat-indentcharwidth-method-word.md)|Indents one or more paragraphs by a specified number of characters.|
|[IndentFirstLineCharWidth](paragraphformat-indentfirstlinecharwidth-method-word.md)|Indents the first line of one or more paragraphs by a specified number of characters.|
|[OpenOrCloseUp](paragraphformat-openorcloseup-method-word.md)|Toggles the spacing before the specified paragraphs.|
|[OpenUp](paragraphformat-openup-method-word.md)|Sets spacing before the specified paragraphs to 12 points.|
|[Reset](paragraphformat-reset-method-word.md)|Removes manual paragraph formatting (formatting not applied using a style).|
|[Space1](paragraphformat-space1-method-word.md)|Single-spaces the specified paragraphs.|
|[Space15](paragraphformat-space15-method-word.md)|Formats the specified paragraphs with 1.5-line spacing.|
|[Space2](paragraphformat-space2-method-word.md)|Double-spaces the specified paragraphs.|
|[TabHangingIndent](paragraphformat-tabhangingindent-method-word.md)|Sets a hanging indent to a specified number of tab stops. .|
|[TabIndent](paragraphformat-tabindent-method-word.md)|Sets the left indent for the specified paragraphs to a specified number of tab stops.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[AddSpaceBetweenFarEastAndAlpha](paragraphformat-addspacebetweenfareastandalpha-property-word.md)| **True** if Microsoft Word is set to automatically add spaces between Japanese and Latin text for the specified paragraphs. This property returns **wdUndefined** if it's set to **True** for only some of the specified paragraphs. Read/write **Long** .|
|[AddSpaceBetweenFarEastAndDigit](paragraphformat-addspacebetweenfareastanddigit-property-word.md)| **True** if Microsoft Word is set to automatically add spaces between Japanese text and numbers for the specified paragraphs. This property returns **wdUndefined** if it's set to **True** for only some of the specified paragraphs. Read/write **Long** .|
|[Alignment](paragraphformat-alignment-property-word.md)|Returns or sets a  **WdParagraphAlignment** constant that represents the alignment for the specified paragraphs. Read/write.|
|[Application](paragraphformat-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[AutoAdjustRightIndent](paragraphformat-autoadjustrightindent-property-word.md)| **True** if Microsoft Word is set to automatically adjust the right indent for the specified paragraphs if you've specified a set number of characters per line. Returns **wdUndefined** if the **AutoAdjustRightIndent** property is set to **True** for only some of the specified paragraphs. Read/write **Long** .|
|[BaseLineAlignment](paragraphformat-baselinealignment-property-word.md)|Returns or sets a  **WdBaselineAlignment** constant that represents the vertical position of fonts on a line. Read/write.|
|[Borders](paragraphformat-borders-property-word.md)|Returns a  **[Borders](borders-object-word.md)** collection that represents all the borders for the specified object.|
|[CharacterUnitFirstLineIndent](paragraphformat-characterunitfirstlineindent-property-word.md)|Returns or sets the value (in characters) for a first-line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent. Read/write  **Single** .|
|[CharacterUnitLeftIndent](paragraphformat-characterunitleftindent-property-word.md)|Returns or sets the left indent value (in characters) for the specified paragraphs. Read/write  **Single** .|
|[CharacterUnitRightIndent](paragraphformat-characterunitrightindent-property-word.md)|Returns or sets the right indent value (in characters) for the specified paragraphs. Read/write  **Single** .|
|[CollapsedByDefault](paragraphformat-collapsedbydefault-property-word.md)|Returns or sets whether the specified paragraph format is collapsed by default. Read-write  **Long**.|
|[Creator](paragraphformat-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[DisableLineHeightGrid](paragraphformat-disablelineheightgrid-property-word.md)| **True** if Microsoft Word aligns characters in the specified paragraphs to the line grid when a set number of lines per page is specified. Returns **wdUndefined** if the **DisableLineHeightGrid** property is set to **True** for only some of the specified paragraphs. Read/write **Long** .|
|[Duplicate](paragraphformat-duplicate-property-word.md)|Returns a read-only  **ParagraphFormat** object that represents the paragraph formatting of the specified paragraph.|
|[FarEastLineBreakControl](paragraphformat-fareastlinebreakcontrol-property-word.md)| **True** if Microsoft Word applies East Asian line-breaking rules to the specified paragraphs. Returns **wdUndefined** if the **FarEastLineBreakControl** property is set to **True** for only some of the specified paragraphs. Read/write **Long** .|
|[FirstLineIndent](paragraphformat-firstlineindent-property-word.md)|Returns or sets the value (in points) for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent. Read/write  **Single** .|
|[HalfWidthPunctuationOnTopOfLine](paragraphformat-halfwidthpunctuationontopofline-property-word.md)| **True** if Microsoft Word changes punctuation symbols at the beginning of a line to half-width characters for the specified paragraphs. This property returns **wdUndefined** if it's set to **True** for only some of the specified paragraphs. Read/write **Long** .|
|[HangingPunctuation](paragraphformat-hangingpunctuation-property-word.md)| **True** if hanging punctuation is enabled for the specified paragraphs. This property returns **wdUndefined** if it's set to **True** for only some of the specified paragraphs. Read/write **Long** .|
|[Hyphenation](paragraphformat-hyphenation-property-word.md)| **True** if the specified paragraphs are included in automatic hyphenation. **False** if the specified paragraphs are to be excluded from automatic hyphenation. Read/write **Long** .|
|[KeepTogether](paragraphformat-keeptogether-property-word.md)| **True** if all lines in the specified paragraphs remain on the same page when Microsoft Word repaginates the document. Read/write **Long** .|
|[KeepWithNext](paragraphformat-keepwithnext-property-word.md)| **True** if the specified paragraph remains on the same page as the paragraph that follows it when Microsoft Word repaginates the document. Read/write **Long** .|
|[LeftIndent](paragraphformat-leftindent-property-word.md)|Returns or sets a  **Single** that represents the left indent value (in points) for the specified paragraph formatting. Read/write.|
|[LineSpacing](paragraphformat-linespacing-property-word.md)|Returns or sets the line spacing (in points) for the specified paragraphs. Read/write  **Single** .|
|[LineSpacingRule](paragraphformat-linespacingrule-property-word.md)|Returns or sets the line spacing for the specified paragraph formatting. Read/write  **[WdLineSpacing](wdlinespacing-enumeration-word.md)** .|
|[LineUnitAfter](paragraphformat-lineunitafter-property-word.md)|Returns or sets the amount of spacing (in gridlines) after the specified paragraphs. Read/write  **Single** .|
|[LineUnitBefore](paragraphformat-lineunitbefore-property-word.md)|Returns or sets the amount of spacing (in gridlines) before the specified paragraphs. Read/write  **Single** .|
|[MirrorIndents](paragraphformat-mirrorindents-property-word.md)|Returns or sets a  **Long** that represents whether left and right indents are the same width. Can be **True** , **False** , or **wdUndefined** . Read/write.|
|[NoLineNumber](paragraphformat-nolinenumber-property-word.md)| **True** if line numbers are repressed for the specified paragraphs. Can be **True** , **False** , or **wdUndefined** . Read/write **Long** .|
|[OutlineLevel](paragraphformat-outlinelevel-property-word.md)|Returns or sets the outline level for the specified paragraphs. Read/write  **[WdOutlineLevel](wdoutlinelevel-enumeration-word.md)** .|
|[PageBreakBefore](paragraphformat-pagebreakbefore-property-word.md)| **True** if a page break is forced before the specified paragraphs. Can be **True** , **False** , or **wdUndefined** . Read/write **Long** .|
|[Parent](paragraphformat-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **ParagraphFormat** object.|
|[ReadingOrder](paragraphformat-readingorder-property-word.md)|Returns or sets the reading order of the specified paragraphs without changing their alignment. Read/write  **WdReadingOrder** .|
|[RightIndent](paragraphformat-rightindent-property-word.md)|Returns or sets the right indent (in points) for the specified paragraphs. Read/write  **Single** .|
|[Shading](paragraphformat-shading-property-word.md)|Returns a  **[Shading](shading-object-word.md)** object that refers to the shading formatting for the specified object.|
|[SpaceAfter](paragraphformat-spaceafter-property-word.md)|Returns or sets the amount of spacing (in points) after the specified paragraph or text column. Read/write  **Single** .|
|[SpaceAfterAuto](paragraphformat-spaceafterauto-property-word.md)| **True** if Microsoft Word automatically sets the amount of spacing after the specified paragraphs. Read/write **Long** .|
|[SpaceBefore](paragraphformat-spacebefore-property-word.md)|Returns or sets the spacing (in points) before the specified paragraphs. Read/write  **Single** .|
|[SpaceBeforeAuto](paragraphformat-spacebeforeauto-property-word.md)| **True** if Microsoft Word automatically sets the amount of spacing before the specified paragraphs. Read/write **Long** .|
|[Style](paragraphformat-style-property-word.md)|Returns or sets the style for the specified object. Read/write  **Variant** .|
|[TabStops](paragraphformat-tabstops-property-word.md)|Returns or sets a  **TabStops** collection that represents all the custom tab stops for the specified paragraphs. Read/write.|
|[TextboxTightWrap](paragraphformat-textboxtightwrap-property-word.md)|Returns or sets a  **[WdTextboxTightWrap](wdtextboxtightwrap-enumeration-word.md)** constant that represents how tightly text wraps around shapes or text boxes. Read/write.|
|[WidowControl](paragraphformat-widowcontrol-property-word.md)| **True** if the first and last lines in the specified paragraph remain on the same page as the rest of the paragraph when Word repaginates the document. Can be **True** , **False** or **wdUndefined** . Read/write **Long** .|
|[WordWrap](paragraphformat-wordwrap-property-word.md)| **True** if Microsoft Word wraps Latin text in the middle of a word in the specified paragraphs or text frames. Read/write **Long** .|

