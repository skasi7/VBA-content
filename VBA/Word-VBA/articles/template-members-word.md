---
title: Template Members (Word)
ms.prod: WORD
ms.assetid: ea133105-b9e9-9169-773d-2c800a88707d
---


# Template Members (Word)
Represents a document template. The  **Template** object is a member of the **[Templates](templates-object-word.md)** collection. The **Templates** collection includes all the available **Template** objects.

Represents a document template. The  **Template** object is a member of the **[Templates](templates-object-word.md)** collection. The **Templates** collection includes all the available **Template** objects.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[OpenAsDocument](template-openasdocument-method-word.md)|Opens the specified template as a document and returns a  **Document** object.|
|[Save](template-save-method-word.md)|Saves the specified template.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](template-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[BuildingBlockEntries](template-buildingblockentries-property-word.md)|Returns a  **[BuildingBlockEntries](buildingblockentries-object-word.md)** collection that represents the collection of building block entries in a template. Read-only.|
|[BuildingBlockTypes](template-buildingblocktypes-property-word.md)|Returns a  **[BuildingBlockTypes](buildingblocktypes-object-word.md)** collection that represents the collection of building block types that are contained in a template. Read-only.|
|[BuiltInDocumentProperties](template-builtindocumentproperties-property-word.md)|Returns a  **DocumentProperties** collection that represents all the built-in document properties for the specified document.|
|[Creator](template-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[CustomDocumentProperties](template-customdocumentproperties-property-word.md)|Returns a  **DocumentProperties** collection that represents all the custom document properties for the specified document.|
|[FarEastLineBreakLanguage](template-fareastlinebreaklanguage-property-word.md)|Returns or sets the East Asian language to use when breaking lines of text in the specified document or template. Read/write  **WdFarEastLineBreakLanguageID** .|
|[FarEastLineBreakLevel](template-fareastlinebreaklevel-property-word.md)|Returns or sets the line break control level for the specified document. Read/write  **WdFarEastLineBreakLevel** .|
|[FullName](template-fullname-property-word.md)|Specifies the name of a template, including the drive or Web path. Read-only  **String** .|
|[JustificationMode](template-justificationmode-property-word.md)|Returns or sets the character spacing adjustment for the specified template. Read/write  **[WdJustificationMode](wdjustificationmode-enumeration-word.md)** .|
|[KerningByAlgorithm](template-kerningbyalgorithm-property-word.md)| **True** if Microsoft Word kerns half-width Latin characters and punctuation marks in the specified document. Read/write **Boolean** .|
|[LanguageID](template-languageid-property-word.md)|Returns or sets a  **[WdLanguageID](wdlanguageid-enumeration-word.md)** constant that represents the language for the specified range. Read/write.|
|[LanguageIDFarEast](template-languageidfareast-property-word.md)|Returns or sets an East Asian language for the specified object. Read/write  **WdLanguageID** .|
|[ListTemplates](template-listtemplates-property-word.md)|Returns a  **[ListTemplates](listtemplates-object-word.md)** collection that represents all the list formats for the specified document, template, or list gallery. Read-only.|
|[Name](template-name-property-word.md)|Returns the name of the specified object. Read-only  **String** .|
|[NoLineBreakAfter](template-nolinebreakafter-property-word.md)|Returns or sets the kinsoku characters after which Microsoft Word will not break a line. Read/write  **String** .|
|[NoLineBreakBefore](template-nolinebreakbefore-property-word.md)|Returns or sets the kinsoku characters before which Microsoft Word will not break a line. Read/write  **String** .|
|[NoProofing](template-noproofing-property-word.md)| **True** if the spelling and grammar checker ignores documents based on this template. Read/write **Long** .|
|[Parent](template-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **Template** object.|
|[Path](template-path-property-word.md)|Returns the path to the specified document template. Read-only  **String** .|
|[Saved](template-saved-property-word.md)| **True** if the specified template has not changed since it was last saved. **False** if Microsoft Word displays a prompt to save changes when the document is closed. Read/write **Boolean** .|
|[Type](template-type-property-word.md)|Returns the template type. Read-only  **[WdTemplateType](wdtemplatetype-enumeration-word.md)** .|
|[VBProject](template-vbproject-property-word.md)|Returns the  **VBProject** object for the specified template.|

