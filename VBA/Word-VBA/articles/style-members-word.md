---
title: Style Members (Word)
ms.prod: WORD
ms.assetid: 37c68e72-c745-bc9c-1547-0cf177cbdef4
---


# Style Members (Word)
Represents a single built-in or user-defined style. The  **Style** object includes style attributes (such as font, font style, and paragraph spacing) as properties of the **Style** object. The **Style** object is a member of the **Styles** collection. The **[Styles](styles-object-word.md)** collection includes all the styles in the specified document.

Represents a single built-in or user-defined style. The  **Style** object includes style attributes (such as font, font style, and paragraph spacing) as properties of the **Style** object. The **Style** object is a member of the **Styles** collection. The **[Styles](styles-object-word.md)** collection includes all the styles in the specified document.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Delete](style-delete-method-word.md)|Deletes the specified style.|
|[LinkToListTemplate](style-linktolisttemplate-method-word.md)|Links the specified style to a list template so that the style's formatting can be applied to lists.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](style-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[AutomaticallyUpdate](style-automaticallyupdate-property-word.md)| **True** if the style is automatically redefined based on the selection. Read/write **Boolean** .|
|[BaseStyle](style-basestyle-property-word.md)|Returns or sets an existing style on which you can base the formatting of another style. Read/write  **Variant** .|
|[Borders](style-borders-property-word.md)|Returns a  **[Borders](borders-object-word.md)** collection that represents all the borders for the specified style.|
|[BuiltIn](style-builtin-property-word.md)| **True** if the specified object is one of the built-in styles or caption labels in Word. Read-only **Boolean** .|
|[Creator](style-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[Description](style-description-property-word.md)|Returns the description of the specified style. Read-only  **String** .|
|[Font](style-font-property-word.md)|Returns or sets a  **[Font](font-object-word.md)** object that represents the character formatting of the specified style. Read/write **Font** .|
|[Frame](style-frame-property-word.md)|Returns a  **[Frame](frame-object-word.md)** object that represents the frame formatting for the specified style. Read-only.|
|[InUse](style-inuse-property-word.md)| **True** if the specified style is a built-in style that has been modified or applied in the document or a new style that has been created in the document. Read-only **Boolean** .|
|[LanguageID](style-languageid-property-word.md)|Returns or sets a  **[WdLanguageID](wdlanguageid-enumeration-word.md)** constant that represents the language for the specified range. Read/write.|
|[LanguageIDFarEast](style-languageidfareast-property-word.md)|Returns or sets an East Asian language for the specified object. Read/write  **[WdLanguageID](wdlanguageid-enumeration-word.md)** .|
|[Linked](style-linked-property-word.md)|Returns or sets a  **Boolean** that represents whether a style is a linked style that can be used for both paragraph and character formatting. Read-only.|
|[LinkStyle](style-linkstyle-property-word.md)|Sets or returns a  **Variant** that represents a link between a paragraph and a character style. Read/write.|
|[ListLevelNumber](style-listlevelnumber-property-word.md)|Returns the list level for the specified style. Read-only  **Long** .|
|[ListTemplate](style-listtemplate-property-word.md)|Returns a  **ListTemplate** object that represents the list formatting for the specified **Style** object.|
|[Locked](style-locked-property-word.md)| **True** if a style cannot be changed or edited. Read/write **Boolean** .|
|[NameLocal](style-namelocal-property-word.md)|Returns the name of a built-in style in the language of the user. Read/write  **String** .|
|[NextParagraphStyle](style-nextparagraphstyle-property-word.md)|Returns or sets the style to be applied automatically to a new paragraph that is inserted after a paragraph formatted with the specified style. Read/write  **Variant** .|
|[NoProofing](style-noproofing-property-word.md)| **True** if the spelling and grammar checker ignores text formatted with this style. Read/write **Long** .|
|[NoSpaceBetweenParagraphsOfSameStyle](style-nospacebetweenparagraphsofsamestyle-property-word.md)| **True** for Microsoft Word to remove spacing between paragraphs that are formatted using the same style. Read/write **Boolean** .|
|[ParagraphFormat](style-paragraphformat-property-word.md)|Returns or sets a  **[ParagraphFormat](paragraphformat-object-word.md)** object that represents the paragraph settings for the specified style. Read/write.|
|[Parent](style-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **Style** object.|
|[Priority](style-priority-property-word.md)|Returns or sets a  **Long** that represents the priority for sorting styles in the **Styles** task pane. Read/write.|
|[QuickStyle](style-quickstyle-property-word.md)|Returns or sets a  **Boolean** that represents whether the style corresponds to an available quick style. Read/write.|
|[Shading](style-shading-property-word.md)|Returns a  **Shading** object that refers to the shading formatting for the specified object.|
|[Table](style-table-property-word.md)|Returns a  **[TableStyle](tablestyle-object-word.md)** object representing properties that can be applied to a table using a table style.|
|[Type](style-type-property-word.md)|Returns the style type. Read-only  **[WdStyleType](wdstyletype-enumeration-word.md)** .|
|[UnhideWhenUsed](style-unhidewhenused-property-word.md)| **True** if the specified style is made visible as a recommended style in the **Styles** and in the **Styles** task pane in Word after it is used in the document. Read/write.|
|[Visibility](style-visibility-property-word.md)| **True** if the specified style is visible as a recommended style in the **Styles** gallery and in the **Styles** task pane. Read/write.|

