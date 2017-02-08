---
title: Find Members (Word)
ms.prod: WORD
ms.assetid: 21f00da0-4c84-ace3-fc79-a55a9ed64360
---


# Find Members (Word)
Represents the criteria for a find operation. 

Represents the criteria for a find operation. 


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[ClearAllFuzzyOptions](find-clearallfuzzyoptions-method-word.md)|Clears all nonspecific search options associated with Japanese text.|
|[ClearFormatting](find-clearformatting-method-word.md)|Removes text and paragraph formatting from the text specified in a find or replace operation.|
|[ClearHitHighlight](find-clearhithighlight-method-word.md)|Removes the highlighting for all text located in a hit highlighting find operation, and returns a  **Boolean** that represents whether the operation was successful.|
|[Execute](find-execute-method-word.md)|Runs the specified find operation. Returns  **True** if the find operation is successful. **Boolean** .|
|[Execute2007](find-execute2007-method-word.md)|Runs the specified find operation. Returns  **True** if the find operation is successful.|
|[HitHighlight](find-hithighlight-method-word.md)|Highlights all found matches and returns a  **Boolean** that represents whether matches were found.|
|[SetAllFuzzyOptions](find-setallfuzzyoptions-method-word.md)|Activates all nonspecific search options associated with Japanese text.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](find-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[CorrectHangulEndings](find-correcthangulendings-property-word.md)| **True** if Microsoft Word automatically corrects Hangul endings when replacing Hangul text. Read/write **Boolean** .|
|[Creator](find-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[Font](find-font-property-word.md)|Returns or sets a  **[Font](font-object-word.md)** object that represents the character formatting of the specified object. Read/write **Font** .|
|[Format](find-format-property-word.md)| **True** if formatting is included in the find operation. Read/write **Boolean** .|
|[Forward](find-forward-property-word.md)| **True** if the find operation searches forward through the document. Read/write **Boolean** .|
|[Found](find-found-property-word.md)| **True** if the search produces a match. Read-only **Boolean** .|
|[Frame](find-frame-property-word.md)|Returns a  **[Frame](frame-object-word.md)** object that represents the frame formatting for the specified style or find-and-replace operation. Read-only.|
|[HanjaPhoneticHangul](find-hanjaphonetichangul-property-word.md)|Returns or sets a  **Boolean** that represents whether to locate phonetic Hangul and hanja characters in a Korean langauge find operation. Read/write.|
|[Highlight](find-highlight-property-word.md)| **True** if highlight formatting is included in the find criteria. Read/write **Long** .|
|[IgnorePunct](find-ignorepunct-property-word.md)|Returns or sets a  **Boolean** that represents whether a find operation should ignore punctuation in found text. Read/write.|
|[IgnoreSpace](find-ignorespace-property-word.md)| Returns or sets a **Boolean** that represents whether a find operation should ignore extra white space in found text. Read/write.|
|[LanguageID](find-languageid-property-word.md)|Returns or sets the language for the specified  **Find** object. Read/write **[WdLanguageID](wdlanguageid-enumeration-word.md)** .|
|[LanguageIDFarEast](find-languageidfareast-property-word.md)|Returns or sets an East Asian language for the specified object. Read/write  **[WdLanguageID](wdlanguageid-enumeration-word.md)** .|
|[LanguageIDOther](find-languageidother-property-word.md)|Returns or sets the language for the specified object. Read/write  **[WdLanguageID](wdlanguageid-enumeration-word.md)** .|
|[MatchAlefHamza](find-matchalefhamza-property-word.md)| **True** if find operations match text with matching alef hamzas in an Arabic language document. Read/write **Boolean** .|
|[MatchAllWordForms](find-matchallwordforms-property-word.md)| **True** if all forms of the text to find are found by the find operation (for instance, if the text to find is "sit," "sat" and "sitting" are found as well). Read/write **Boolean** .|
|[MatchByte](find-matchbyte-property-word.md)| **True** if Microsoft Word distinguishes between full-width and half-width letters or characters during a search. Read/write **Boolean** .|
|[MatchCase](find-matchcase-property-word.md)| **True** if the find operation is case sensitive. The default is **False** . Read/write **Boolean** .|
|[MatchControl](find-matchcontrol-property-word.md)| **True** if find operations match text with matching bidirectional control characters in a right-to-left language document. Read/write **Boolean** .|
|[MatchDiacritics](find-matchdiacritics-property-word.md)| **True** if find operations match text with matching diacritics in a right-to-left language document. Read/write **Boolean** .|
|[MatchFuzzy](find-matchfuzzy-property-word.md)| **True** if Microsoft Word uses the nonspecific search options for Japanese text during a search. Read/write **Boolean** .|
|[MatchKashida](find-matchkashida-property-word.md)| **True** if find operations match text with matching kashidas in an Arabic language document. Read/write **Boolean** .|
|[MatchPhrase](find-matchphrase-property-word.md)| **True** ignores all white space and control characters between words. Read/write.|
|[MatchPrefix](find-matchprefix-property-word.md)| **True** to match words beginning with the search string. Read/write.|
|[MatchSoundsLike](find-matchsoundslike-property-word.md)| **True** if words that sound similar to the text to find are returned by the find operation. Read/write **Boolean** .|
|[MatchSuffix](find-matchsuffix-property-word.md)| **True** to match words ending with the search string. Read/write.|
|[MatchWholeWord](find-matchwholeword-property-word.md)| **True** if the find operation locates only entire words and not text that's part of a larger word. Read/write **Boolean** .|
|[MatchWildcards](find-matchwildcards-property-word.md)| **True** if the text to find contains wildcards. Read/write **Boolean** .|
|[NoProofing](find-noproofing-property-word.md)| **True** if Microsoft Word finds or replaces text that the spelling and grammar checker ignores. Read/write **Long** .|
|[ParagraphFormat](find-paragraphformat-property-word.md)|Returns or sets a  **[ParagraphFormat](paragraphformat-object-word.md)** object that represents the paragraph settings for the specified find operation. Read/write.|
|[Parent](find-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **Find** object.|
|[Replacement](find-replacement-property-word.md)|Returns a  **[Replacement](replacement-object-word.md)** object that contains the criteria for a replace operation.|
|[Style](find-style-property-word.md)|Returns or sets the style for the specified object. Read/write  **Variant** .|
|[Text](find-text-property-word.md)|Returns or sets the text to find. Read/write  **String** .|
|[Wrap](find-wrap-property-word.md)|Returns or sets what happens if the search begins at a point other than the beginning of the document and the end of the document is reached (or vice versa if  **Forward** is set to **False** ) or if the search text isn't found in the specified selection or range. Read/write **WdFindWrap** .|

