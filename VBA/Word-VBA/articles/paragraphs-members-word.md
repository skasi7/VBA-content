---
title: Paragraphs Members (Word)
ms.prod: WORD
ms.assetid: 490e2695-3cdd-4906-f730-583d18486aa2
---


# Paragraphs Members (Word)
A collection of  **[Paragraph](paragraph-object-word.md)** objects in a selection, range, or document.

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Add](paragraphs-add-method-word.md)|Returns a  **Paragraph** object that represents a new, blank paragraph added to a document.|
|[CloseUp](paragraphs-closeup-method-word.md)|Removes any spacing before the specified paragraphs.|
|[DecreaseSpacing](paragraphs-decreasespacing-method-word.md)|Decreases the spacing before and after paragraphs in six-point increments.|
|[IncreaseSpacing](paragraphs-increasespacing-method-word.md)|Increases the spacing before and after paragraphs in six-point increments.|
|[Indent](paragraphs-indent-method-word.md)|Indents one or more paragraphs by one level.|
|[IndentCharWidth](paragraphs-indentcharwidth-method-word.md)|Indents one or more paragraphs by a specified number of characters.|
|[IndentFirstLineCharWidth](paragraphs-indentfirstlinecharwidth-method-word.md)|Indents the first line of one or more paragraphs by a specified number of characters.|
|[Item](paragraphs-item-method-word.md)|Returns an individual  **Paragraph** object in a collection.|
|[OpenOrCloseUp](paragraphs-openorcloseup-method-word.md)|Toggles spacing before paragraphs.|
|[OpenUp](paragraphs-openup-method-word.md)|Sets spacing before the specified paragraphs to 12 points.|
|[Outdent](paragraphs-outdent-method-word.md)|Removes one level of indent for one or more paragraphs.|
|[OutlineDemote](paragraphs-outlinedemote-method-word.md)|Applies the next heading level style (Heading 1 through Heading 8) to the specified paragraphs.|
|[OutlineDemoteToBody](paragraphs-outlinedemotetobody-method-word.md)|Demotes the specified paragraph or paragraphs to body text by applying the Normal style.|
|[OutlinePromote](paragraphs-outlinepromote-method-word.md)|Applies the previous heading level style (Heading 1 through Heading 8) to the specified paragraph or paragraphs.|
|[Reset](paragraphs-reset-method-word.md)|Removes manual paragraph formatting (formatting not applied using a style). .|
|[Space1](paragraphs-space1-method-word.md)|Single-spaces the specified paragraphs.|
|[Space15](paragraphs-space15-method-word.md)|Formats the specified paragraphs with 1.5-line spacing.|
|[Space2](paragraphs-space2-method-word.md)|Double-spaces the specified paragraphs. .|
|[TabHangingIndent](paragraphs-tabhangingindent-method-word.md)|Sets a hanging indent to a specified number of tab stops.|
|[TabIndent](paragraphs-tabindent-method-word.md)|Sets the left indent for the specified paragraphs to a specified number of tab stops.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[AddSpaceBetweenFarEastAndAlpha](paragraphs-addspacebetweenfareastandalpha-property-word.md)| **True** if Microsoft Word is set to automatically add spaces between Japanese and Latin text for the specified paragraphs. This property returns **wdUndefined** if it's set to **True** for only some of the specified paragraphs. Read/write **Long** .|
|[AddSpaceBetweenFarEastAndDigit](paragraphs-addspacebetweenfareastanddigit-property-word.md)| **True** if Microsoft Word is set to automatically add spaces between Japanese text and numbers for the specified paragraphs. This property returns **wdUndefined** if it's set to **True** for only some of the specified paragraphs. Read/write **Long** .|
|[Alignment](paragraphs-alignment-property-word.md)|Returns or sets a  **WdParagraphAlignment** constant that represents the alignment for the specified paragraphs. Read/write.|
|[Application](paragraphs-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[AutoAdjustRightIndent](paragraphs-autoadjustrightindent-property-word.md)| **True** if Microsoft Word is set to automatically adjust the right indent for the specified paragraphs if you've specified a set number of characters per line. Returns **wdUndefined** if the **AutoAdjustRightIndent** property is set to **True** for only some of the specified paragraphs. Read/write **Long** .|
|[BaseLineAlignment](paragraphs-baselinealignment-property-word.md)|Returns or sets a  **WdBaselineAlignment** constant that represents the vertical position of fonts on a line. Read/write.|
|[Borders](paragraphs-borders-property-word.md)|Returns a  **[Borders](borders-object-word.md)** collection that represents all the borders for the specified object.|
|[CharacterUnitFirstLineIndent](paragraphs-characterunitfirstlineindent-property-word.md)|Returns or sets the value (in characters) for a first-line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent. Read/write  **Single** .|
|[CharacterUnitLeftIndent](paragraphs-characterunitleftindent-property-word.md)|Returns or sets the left indent value (in characters) for the specified paragraphs. Read/write  **Single** .|
|[CharacterUnitRightIndent](paragraphs-characterunitrightindent-property-word.md)|Returns or sets the right indent value (in characters) for the specified paragraphs. Read/write  **Single** .|
|[Count](paragraphs-count-property-word.md)|Returns a  **Long** that represents the number of paragraphs in the collection. Read-only.|
|[Creator](paragraphs-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[DisableLineHeightGrid](paragraphs-disablelineheightgrid-property-word.md)| **True** if Microsoft Word aligns characters in the specified paragraphs to the line grid when a set number of lines per page is specified. Returns **wdUndefined** if the **DisableLineHeightGrid** property is set to **True** for only some of the specified paragraphs. Read/write **Long** .|
|[FarEastLineBreakControl](paragraphs-fareastlinebreakcontrol-property-word.md)| **True** if Microsoft Word applies East Asian line-breaking rules to the specified paragraphs. Returns **wdUndefined** if the **FarEastLineBreakControl** property is set to **True** for only some of the specified paragraphs. Read/write **Long** .|
|[First](paragraphs-first-property-word.md)|Returns a  **[Paragraph](paragraph-object-word.md)** object that represents the first item in the **Paragraphs** collection.|
|[FirstLineIndent](paragraphs-firstlineindent-property-word.md)|Returns or sets the value (in points) for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent. Read/write  **Single** .|
|[Format](paragraphs-format-property-word.md)|Returns or sets a  **[ParagraphFormat](paragraphformat-object-word.md)** object that represents the formatting of the specified paragraph or paragraphs.|
|[HalfWidthPunctuationOnTopOfLine](paragraphs-halfwidthpunctuationontopofline-property-word.md)| **True** if Microsoft Word changes punctuation symbols at the beginning of a line to half-width characters for the specified paragraphs. This property returns **wdUndefined** if it's set to **True** for only some of the specified paragraphs. Read/write **Long** .|
|[HangingPunctuation](paragraphs-hangingpunctuation-property-word.md)| **True** if hanging punctuation is enabled for the specified paragraphs. This property returns **wdUndefined** if it's set to **True** for only some of the specified paragraphs. Read/write **Long** .|
|[Hyphenation](paragraphs-hyphenation-property-word.md)| **True** if the specified paragraphs are included in automatic hyphenation. **False** if the specified paragraphs are to be excluded from automatic hyphenation. Read/write **Long** .|
|[KeepTogether](paragraphs-keeptogether-property-word.md)| **True** if all lines in the specified paragraphs remain on the same page when Microsoft Word repaginates the document. Read/write **Long** .|
|[KeepWithNext](paragraphs-keepwithnext-property-word.md)| **True** if the specified paragraphs remain on the same page as the paragraphs that follow it when Microsoft Word repaginates the document. Read/write **Long** .|
|[Last](paragraphs-last-property-word.md)|Returns a  **Paragraph** object that represents the last item in the collection of paragraphs.|
|[LeftIndent](paragraphs-leftindent-property-word.md)|Returns or sets a  **Single** that represents the left indent value (in points) for the specified paragraphs. Read/write.|
|[LineSpacing](paragraphs-linespacing-property-word.md)|Returns or sets the line spacing (in points) for the specified paragraphs. Read/write  **Single** .|
|[LineSpacingRule](paragraphs-linespacingrule-property-word.md)|Returns or sets the line spacing for the specified paragraphs. Read/write  **WdLineSpacing** .|
|[LineUnitAfter](paragraphs-lineunitafter-property-word.md)|Returns or sets the amount of spacing (in gridlines) after the specified paragraphs. Read/write  **Single** .|
|[LineUnitBefore](paragraphs-lineunitbefore-property-word.md)|Returns or sets the amount of spacing (in gridlines) before the specified paragraphs. Read/write  **Single** .|
|[NoLineNumber](paragraphs-nolinenumber-property-word.md)| **True** if line numbers are repressed for the specified paragraphs. Can be **True** , **False** , or **wdUndefined** . Read/write **Long** .|
|[OutlineLevel](paragraphs-outlinelevel-property-word.md)|Returns or sets the outline level for the specified paragraphs. Read/write  **[WdOutlineLevel](wdoutlinelevel-enumeration-word.md)** .|
|[PageBreakBefore](paragraphs-pagebreakbefore-property-word.md)| **True** if a page break is forced before the specified paragraphs. Can be **True** , **False** , or **wdUndefined** . Read/write **Long** .|
|[Parent](paragraphs-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **Paragraphs** object.|
|[ReadingOrder](paragraphs-readingorder-property-word.md)|Returns or sets the reading order of the specified paragraphs without changing their alignment. Read/write  **WdReadingOrder** .|
|[RightIndent](paragraphs-rightindent-property-word.md)|Returns or sets the right indent (in points) for the specified paragraphs. Read/write  **Single** .|
|[Shading](paragraphs-shading-property-word.md)|Returns a  **Shading** object that refers to the shading formatting for the specified paragraphs.|
|[SpaceAfter](paragraphs-spaceafter-property-word.md)|Returns or sets the amount of spacing (in points) after the specified paragraph or text column. Read/write  **Single** .|
|[SpaceAfterAuto](paragraphs-spaceafterauto-property-word.md)| **True** if Microsoft Word automatically sets the amount of spacing after the specified paragraphs. Read/write **Long** .|
|[SpaceBefore](paragraphs-spacebefore-property-word.md)|Returns or sets the spacing (in points) before the specified paragraphs. Read/write  **Single** .|
|[SpaceBeforeAuto](paragraphs-spacebeforeauto-property-word.md)| **True** if Microsoft Word automatically sets the amount of spacing before the specified paragraphs. Read/write **Long** .|
|[Style](paragraphs-style-property-word.md)|Returns or sets the style for the specified paragraphs. Read/write  **Variant** .|
|[TabStops](paragraphs-tabstops-property-word.md)|Returns or sets a  **TabStops** collection that represents all the custom tab stops for the specified paragraphs. Read/write.|
|[WidowControl](paragraphs-widowcontrol-property-word.md)| **True** if the first and last lines in the specified paragraph remain on the same page as the rest of the paragraph when Word repaginates the document. Can be **True** , **False** or **wdUndefined** . Read/write **Long** .|
|[WordWrap](paragraphs-wordwrap-property-word.md)| **True** if Microsoft Word wraps Latin text in the middle of a word in the specified paragraphs. Read/write **Long** .|

