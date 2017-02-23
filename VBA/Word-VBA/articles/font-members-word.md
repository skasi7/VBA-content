---
title: Font Members (Word)
ms.prod: WORD
ms.assetid: 04a3c706-4062-09bc-70d9-cef3748a7d57
---


# Font Members (Word)
Contains font attributes (such as font name, font size and color) for an object.

Contains font attributes (such as font name, font size and color) for an object.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Grow](font-grow-method-word.md)|Increases the font size to the next available size.|
|[Reset](font-reset-method-word.md)|Removes manual character formatting (formatting not applied using a style). For example, if you manually format a word as bold and the underlying style is plain text (not bold), the  **Reset** method removes the bold format.|
|[SetAsTemplateDefault](font-setastemplatedefault-method-word.md)|Sets the specified font formatting as the default for the active document and all new documents based on the active template.|
|[Shrink](font-shrink-method-word.md)|Decreases the font size to the next available size.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[AllCaps](font-allcaps-property-word.md)| **True** if the font is formatted as all capital letters. Read/write **Long** .|
|[Application](font-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[Bold](font-bold-property-word.md)| **True** if the font is formatted as bold. Read/write **Long** .|
|[BoldBi](font-boldbi-property-word.md)| **True** if the font is formatted as bold. Read/write **Long** .|
|[Borders](font-borders-property-word.md)|Returns a  **[Borders](borders-object-word.md)** collection that represents all the borders for the specified font.|
|[ColorIndex](font-colorindex-property-word.md)|Returns or sets a  **WdColorIndex** constant that represents the color for the specified font. Read/write .|
|[ColorIndexBi](font-colorindexbi-property-word.md)|Returns or sets the color for the specified  **Font** object in a right-to-left language document. Read/write **WdColorIndex** .|
|[ContextualAlternates](font-contextualalternates-property-word.md)|Specifies whether or not contextual alternates are enabled for the specified font. Read/write  **Long** .|
|[Creator](font-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[DiacriticColor](font-diacriticcolor-property-word.md)|Returns or sets the 24-bit color to be used for diacritics for the specified  **Font** object. Read/write.|
|[DisableCharacterSpaceGrid](font-disablecharacterspacegrid-property-word.md)| **True** if Microsoft Word ignores the number of characters per line for the corresponding **Font** object. Read/write **Boolean** .|
|[DoubleStrikeThrough](font-doublestrikethrough-property-word.md)| **True** if the specified font is formatted as double strikethrough text. .|
|[Duplicate](font-duplicate-property-word.md)|Returns a read-only  **Font** object that represents the character formatting of the specified font.|
|[Emboss](font-emboss-property-word.md)| **True** if the specified font is formatted as embossed. Read/write **Long** .|
|[EmphasisMark](font-emphasismark-property-word.md)|Returns or sets a  **WdEmphasisMark** constant that represents the emphasis mark for a character or designated character string. Read/write.|
|[Engrave](font-engrave-property-word.md)| **True** if the font is formatted as engraved. Read/write **Long** .|
|[Fill](font-fill-property-word.md)|Returns a [FillFormat](fillformat-object-word.md) object that contains fill formatting properties for the font used by the specified range of text. Read-only.|
|[Glow](font-glow-property-word.md)|Returns a [GlowFormat](glowformat-object-word.md) object that represents the glow formatting for the font used by the specified range of text. Read-only.|
|[Hidden](font-hidden-property-word.md)| **True** if the font is formatted as hidden text. Read/write **Long** .|
|[Italic](font-italic-property-word.md)| **True** if the font or range is formatted as italic. Read/write **Long** .|
|[ItalicBi](font-italicbi-property-word.md)| **True** if the font or range is formatted as italic. Read/write **Long** .|
|[Kerning](font-kerning-property-word.md)|Returns or sets the minimum font size for which Microsoft Word will adjust kerning automatically. Read/write  **Single** .|
|[Ligatures](font-ligatures-property-word.md)|Returns or sets the ligatures setting for the specified  **Font** object. Read/write[WdLigatures](wdligatures-enumeration-word.md).|
|[Line](font-line-property-word.md)|Returns a [LineFormat](lineformat-object-word.md) object that specifies the formatting for a line. Read/write.|
|[Name](font-name-property-word.md)|Returns or sets the name of the specified object. Read/write  **String** .|
|[NameAscii](font-nameascii-property-word.md)|Returns or sets the font used for Latin text (characters with character codes from 0 (zero) through 127). Read/write  **String** .|
|[NameBi](font-namebi-property-word.md)|Returns or sets the name of the font in a right-to-left language document. Read/write  **String** .|
|[NameFarEast](font-namefareast-property-word.md)|Returns or sets an East Asian font name. Read/write  **String** .|
|[NameOther](font-nameother-property-word.md)|Returns or sets the font used for characters with character codes from 128 through 255. Read/write  **String** .|
|[NumberForm](font-numberform-property-word.md)|Returns or sets the number form setting for an OpenType font. Read/write [WdNumberForm](wdnumberform-enumeration-word.md).|
|[NumberSpacing](font-numberspacing-property-word.md)|Returns or sets the number spacing setting for a font. Read/write [WdNumberSpacing](wdnumberspacing-enumeration-word.md).|
|[Outline](font-outline-property-word.md)| **True** if the font is formatted as outline. Read/write **Long** .|
|[Parent](font-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **Font** object.|
|[Position](font-position-property-word.md)|Returns or sets the position of text (in points) relative to the base line. Read/write  **Long** .|
|[Reflection](font-reflection-property-word.md)|Returns a [ReflectionFormat](reflectionformat-object-word.md) object that represents the reflection formatting for a shape. Read-only.|
|[Scaling](font-scaling-property-word.md)|Returns or sets the scaling percentage applied to the font. Read/write  **Long** .|
|[Shading](font-shading-property-word.md)|Returns a  **Shading** object that refers to the shading formatting for the specified font.|
|[Shadow](font-shadow-property-word.md)| **True** if the specified font is formatted as shadowed. Read/write **Long** .|
|[Size](font-size-property-word.md)|Returns or sets the font size, in points. Read/write  **Single** .|
|[SizeBi](font-sizebi-property-word.md)|Returns or sets the font size in points. Read/write  **Single** .|
|[SmallCaps](font-smallcaps-property-word.md)| **True** if the font is formatted as small capital letters. Read/write **Long** .|
|[Spacing](font-spacing-property-word.md)|Returns or sets the spacing (in points) between characters. Read/write  **Single** .|
|[StrikeThrough](font-strikethrough-property-word.md)| **True** if the font is formatted as strikethrough text. Read/write **Long** .|
|[StylisticSet](font-stylisticset-property-word.md)|Specifies the stylistic set for the specified font. Read/write [WdStylisticSet](wdstylisticset-enumeration-word.md).|
|[Subscript](font-subscript-property-word.md)| **True** if the font is formatted as subscript. Read/write **Long** .|
|[Superscript](font-superscript-property-word.md)| **True** if the font is formatted as superscript. Read/write **Long** .|
|[TextColor](font-textcolor-property-word.md)|Returns a [ColorFormat](colorformat-object-word.md) object that represents the color for the specified font. Read-only.|
|[TextShadow](font-textshadow-property-word.md)|Returns a [ShadowFormat](shadowformat-object-word.md) object that specifies the shadow formatting for the specified font.|
|[ThreeD](font-threed-property-word.md)|Returns a [ThreeDFormat](threedformat-object-word.md) object that contains 3-D effect formatting properties for the specified font. Read-only.|
|[Underline](font-underline-property-word.md)|Returns or sets the type of underline applied to the font. Read/write  **[WdUnderline](wdunderline-enumeration-word.md)** .|
|[UnderlineColor](font-underlinecolor-property-word.md)|Returns or sets the 24-bit color of the underline for the specified  **Font** object. .|

