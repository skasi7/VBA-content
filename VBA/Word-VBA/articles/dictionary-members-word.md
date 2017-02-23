---
title: Dictionary Members (Word)
ms.prod: WORD
ms.assetid: 40366ef7-9a5e-19f5-088f-00b36bec68f4
---


# Dictionary Members (Word)
Represents a dictionary.  **Dictionary** objects that represent custom dictionaries are members of the **[Dictionaries](dictionaries-object-word.md)** collection. Other dictionary objects are returned by properties of the **[Languages](languages-object-word.md)** collection; these include the **[ActiveSpellingDictionary](language-activespellingdictionary-property-word.md)** , **[ActiveGrammarDictionary](language-activegrammardictionary-property-word.md)** , **[ActiveThesaurusDictionary](language-activethesaurusdictionary-property-word.md)** , and **[ActiveHyphenationDictionary](language-activehyphenationdictionary-property-word.md)** properties.

Represents a dictionary.  **Dictionary** objects that represent custom dictionaries are members of the **[Dictionaries](dictionaries-object-word.md)** collection. Other dictionary objects are returned by properties of the **[Languages](languages-object-word.md)** collection; these include the **[ActiveSpellingDictionary](language-activespellingdictionary-property-word.md)** , **[ActiveGrammarDictionary](language-activegrammardictionary-property-word.md)** , **[ActiveThesaurusDictionary](language-activethesaurusdictionary-property-word.md)** , and **[ActiveHyphenationDictionary](language-activehyphenationdictionary-property-word.md)** properties.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Delete](dictionary-delete-method-word.md)|Deletes the specified dictionary.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](dictionary-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[Creator](dictionary-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[LanguageID](dictionary-languageid-property-word.md)|Returns or sets a  **[WdLanguageID](wdlanguageid-enumeration-word.md)** constant that represents the language for the specified object. Read/write.|
|[LanguageSpecific](dictionary-languagespecific-property-word.md)| **True** if the custom dictionary is to be used only with text formatted for a specific language. Read/write **Boolean** .|
|[Name](dictionary-name-property-word.md)|Returns the name of the specified object. Read-only  **String** .|
|[Parent](dictionary-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **Dictionary** object.|
|[Path](dictionary-path-property-word.md)|Returns the path to the specified dictionary. Read-only  **String** .|
|[ReadOnly](dictionary-readonly-property-word.md)| **True** if the specified dictionary cannot be changed. Read-only **Boolean** .|
|[Type](dictionary-type-property-word.md)|Returns the dictionary type. Read-only  **WdDictionaryType** .|

