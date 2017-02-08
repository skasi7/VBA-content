---
title: Replacement Members (Word)
ms.prod: WORD
ms.assetid: 013ead94-f79c-fc4f-164b-49b2a88b3e88
---


# Replacement Members (Word)
Represents the replace criteria for a find-and-replace operation. The properties and methods of the  **Replacement** object correspond to the options in the **Find and Replace** dialog box.

Represents the replace criteria for a find-and-replace operation. The properties and methods of the  **Replacement** object correspond to the options in the **Find and Replace** dialog box.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[ClearFormatting](replacement-clearformatting-method-word.md)|Removes text and paragraph formatting from the text specified in a replace operation.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](replacement-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[Creator](replacement-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[Font](replacement-font-property-word.md)|Returns or sets a  **[Font](font-object-word.md)** object that represents the character formatting of the specified object. Read/write **Font** .|
|[Frame](replacement-frame-property-word.md)|Returns a  **[Frame](frame-object-word.md)** object that represents the frame formatting for the specified style or find-and-replace operation. Read-only.|
|[Highlight](replacement-highlight-property-word.md)| **True** if highlight formatting is applied to the replacement text. Read/write **Long** .|
|[LanguageID](replacement-languageid-property-word.md)|Returns or sets a  **[WdLanguageID](wdlanguageid-enumeration-word.md)** constant that represents the language for the specified range. Read/write.|
|[LanguageIDFarEast](replacement-languageidfareast-property-word.md)|Returns or sets an East Asian language for the specified replacement. Read/write  **[WdLanguageID](wdlanguageid-enumeration-word.md)** .|
|[NoProofing](replacement-noproofing-property-word.md)| **True** if Microsoft Word replaces text that the spelling and grammar checker ignores. Read/write **Long** .|
|[ParagraphFormat](replacement-paragraphformat-property-word.md)|Returns or sets a  **[ParagraphFormat](paragraphformat-object-word.md)** object that represents the paragraph settings for the specified replacement operation. Read/write.|
|[Parent](replacement-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **Replacement** object.|
|[Style](replacement-style-property-word.md)|Returns or sets the style for the specified object. To set this property, specify the local name of the style, an integer, a  **[WdBuiltinStyle](wdbuiltinstyle-enumeration-word.md)** constant, or an object that represents the style. Read/write **Variant** .|
|[Text](replacement-text-property-word.md)|Returns or sets the text to replace. Read/write  **String** .|

