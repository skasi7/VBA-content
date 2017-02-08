---
title: EmailOptions Members (Word)
ms.prod: WORD
ms.assetid: 0f8a549b-283c-dc9d-dc1e-1179a9d6fb0b
---


# EmailOptions Members (Word)
Contains global application-level attributes used by Microsoft Word when you create and edit e-mail messages and replies.

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](emailoptions-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[AutoFormatAsYouTypeApplyBorders](emailoptions-autoformatasyoutypeapplyborders-property-word.md)| **True** if a series of three or more hyphens (-), equal signs (=), or underscore characters (_) are automatically replaced by a specific border line when the ENTER key is pressed. Read/write **Boolean** .|
|[AutoFormatAsYouTypeApplyBulletedLists](emailoptions-autoformatasyoutypeapplybulletedlists-property-word.md)| **True** if bullet characters (such as asterisks, hyphens, and greater-than signs) are replaced with bullets. Read/write **Boolean** .|
|[AutoFormatAsYouTypeApplyClosings](emailoptions-autoformatasyoutypeapplyclosings-property-word.md)| **True** for Microsoft Word to automatically apply the Closing style to letter closings as you type. Read/write **Boolean** .|
|[AutoFormatAsYouTypeApplyDates](emailoptions-autoformatasyoutypeapplydates-property-word.md)| **True** for Microsoft Word to automatically apply the Date style to dates as you type. Read/write.|
|[AutoFormatAsYouTypeApplyFirstIndents](emailoptions-autoformatasyoutypeapplyfirstindents-property-word.md)| **True** for Microsoft Word to automatically replace a space entered at the beginning of a paragraph with a first-line indent. Read/write.|
|[AutoFormatAsYouTypeApplyHeadings](emailoptions-autoformatasyoutypeapplyheadings-property-word.md)| **True** if styles are automatically applied to headings as you type. Read/write **Boolean** .|
|[AutoFormatAsYouTypeApplyNumberedLists](emailoptions-autoformatasyoutypeapplynumberedlists-property-word.md)| **True** if paragraphs are automatically formatted as numbered lists. Read/write **Boolean** .|
|[AutoFormatAsYouTypeApplyTables](emailoptions-autoformatasyoutypeapplytables-property-word.md)| **True** if Word automatically creates a table when you type a plus sign, a series of hyphens, another plus sign, and so on, and then press ENTER. Read/write **Boolean** .|
|[AutoFormatAsYouTypeAutoLetterWizard](emailoptions-autoformatasyoutypeautoletterwizard-property-word.md)| **True** for Microsoft Word to automatically start the Letter Wizard when the user enters a letter salutation or closing. Read/write.|
|[AutoFormatAsYouTypeDefineStyles](emailoptions-autoformatasyoutypedefinestyles-property-word.md)| **True** if Word automatically creates new styles based on manual formatting. Read/write **Boolean** .|
|[AutoFormatAsYouTypeDeleteAutoSpaces](emailoptions-autoformatasyoutypedeleteautospaces-property-word.md)| **True** for Microsoft Word to automatically delete spaces inserted between Japanese and Latin text as you type. Read/write.|
|[AutoFormatAsYouTypeFormatListItemBeginning](emailoptions-autoformatasyoutypeformatlistitembeginning-property-word.md)| **True** if Word repeats character formatting applied to the beginning of a list item to the next list item. Read/write **Boolean** .|
|[AutoFormatAsYouTypeInsertClosings](emailoptions-autoformatasyoutypeinsertclosings-property-word.md)| **True** for Microsoft Word to automatically insert the corresponding memo closing when the user enters a memo heading. Read/write.|
|[AutoFormatAsYouTypeInsertOvers](emailoptions-autoformatasyoutypeinsertovers-property-word.md)| **True** for Microsoft Word to automatically insert "以上" when the user enters "記" or "案". Read/write **Boolean** .|
|[AutoFormatAsYouTypeMatchParentheses](emailoptions-autoformatasyoutypematchparentheses-property-word.md)| **True** for Microsoft Word to automatically correct improperly paired parentheses. Read/write.|
|[AutoFormatAsYouTypeReplaceFarEastDashes](emailoptions-autoformatasyoutypereplacefareastdashes-property-word.md)| **True** for Microsoft Word to automatically correct long vowel sounds and dashes. Read/write.|
|[AutoFormatAsYouTypeReplaceFractions](emailoptions-autoformatasyoutypereplacefractions-property-word.md)| **True** if typed fractions are replaced with fractions from the current character set as you type; for example, "1/2" is replaced with "½." Read/write **Boolean** .|
|[AutoFormatAsYouTypeReplaceHyperlinks](emailoptions-autoformatasyoutypereplacehyperlinks-property-word.md)| **True** if e-mail addresses, server and share names (also known as UNC paths), and Internet addresses (also known as URLs) are automatically changed to hyperlinks as you type. Read/write **Boolean** .|
|[AutoFormatAsYouTypeReplaceOrdinals](emailoptions-autoformatasyoutypereplaceordinals-property-word.md)| **True** if the ordinal number suffixes "st", "nd", "rd", and "th" are replaced with the same letters in superscript as you type; for example, "1st" is replaced with "1" followed by "st" formatted as superscript. Read/write **Boolean** .|
|[AutoFormatAsYouTypeReplacePlainTextEmphasis](emailoptions-autoformatasyoutypereplaceplaintextemphasis-property-word.md)| **True** if manual emphasis characters are automatically replaced with character formatting as you type; for example, "*bold*" is changed to " **bold** ". Read/write **Boolean** .|
|[AutoFormatAsYouTypeReplaceQuotes](emailoptions-autoformatasyoutypereplacequotes-property-word.md)| **True** if straight quotation marks are automatically changed to smart (curly) quotation marks as you type. Read/write **Boolean** .|
|[AutoFormatAsYouTypeReplaceSymbols](emailoptions-autoformatasyoutypereplacesymbols-property-word.md)| **True** if two consecutive hyphens (--) are replaced with an en dash (-) or an em dash (—) as you type. Read/write **Boolean** .|
|[ComposeStyle](emailoptions-composestyle-property-word.md)|Returns a  **[Style](style-object-word.md)** object that represents the style used to compose new e-mail messages. Read-only.|
|[Creator](emailoptions-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[EmailSignature](emailoptions-emailsignature-property-word.md)|Returns an  **[EmailSignature](emailsignature-object-word.md)** object that represents the signatures Microsoft Word appends to outgoing e-mail messages. Read-only.|
|[HTMLFidelity](emailoptions-htmlfidelity-property-word.md)|Strips HTML tags used for opening HTML files in Word but not required for display. Read/write  **WdEmailHTMLFidelity** .|
|[MarkComments](emailoptions-markcomments-property-word.md)| **True** if Microsoft Word marks the user's comments in e-mail messages. Read/write **Boolean** .|
|[MarkCommentsWith](emailoptions-markcommentswith-property-word.md)|Returns or sets the string with which Microsoft Word marks comments in e-mail messages. Read/write  **String** .|
|[NewColorOnReply](emailoptions-newcoloronreply-property-word.md)| **True** specifies whether a user needs to choose a new color for reply text when replying to e-mail. Read/write **Boolean** .|
|[Parent](emailoptions-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **EmailOptions** object.|
|[PlainTextStyle](emailoptions-plaintextstyle-property-word.md)|Returns the  **[Style](style-object-word.md)** object that represents the text attributes for e-mail messages that are sent or received using plain text.|
|[RelyOnCSS](emailoptions-relyoncss-property-word.md)| **True** if cascading style sheets (CSS) are used for font formatting when you view a saved document in a Web browser. Read/write **Boolean** .|
|[ReplyStyle](emailoptions-replystyle-property-word.md)|Returns a  **[Style](style-object-word.md)** object that represents the style used when replying to e-mail messages.|
|[TabIndentKey](emailoptions-tabindentkey-property-word.md)| **True** if the TAB and BACKSPACE keys can be used to increase and decrease, respectively, the left indent of paragraphs and if the BACKSPACE key can be used to change right-aligned paragraphs to centered paragraphs and centered paragraphs to left-aligned paragraphs. Read/write **Boolean** .|
|[ThemeName](emailoptions-themename-property-word.md)|Returns or sets the name of the theme plus any theme formatting options to use for new e-mail messages. Read/write  **String** .|
|[UseThemeStyle](emailoptions-usethemestyle-property-word.md)| **True** if new e-mail messages use the character style defined by the default e-mail message theme. Read/write **Boolean** .|
|[UseThemeStyleOnReply](emailoptions-usethemestyleonreply-property-word.md)| **True** for Microsoft Word to use a theme when replying to e-mail. Read/write **Boolean** .|

