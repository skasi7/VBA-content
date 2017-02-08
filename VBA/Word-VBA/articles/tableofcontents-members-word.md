---
title: TableOfContents Members (Word)
ms.prod: WORD
ms.assetid: bfd1b65b-98c3-a60b-6668-34dd05f6ee85
---


# TableOfContents Members (Word)
Represents a single table of contents in a document. The  **TableOfContents** object is a member of the **[TablesOfContents](tablesofcontents-object-word.md)** collection. The **TablesOfContents** collection includes all the tables of contents in a document.

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Delete](tableofcontents-delete-method-word.md)|Deletes the specified table of contents.|
|[Update](tableofcontents-update-method-word.md)|Updates the entries shown in a table of contents.|
|[UpdatePageNumbers](tableofcontents-updatepagenumbers-method-word.md)|Updates the page numbers for items in the specified table of contents.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](tableofcontents-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[Creator](tableofcontents-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[HeadingStyles](tableofcontents-headingstyles-property-word.md)|Returns a  **[HeadingStyles](headingstyles-object-word.md)** object that represents additional styles used to compile a table of contents or table of figures (styles other than the Heading 1 - Heading 9 styles). Read-only.|
|[HidePageNumbersInWeb](tableofcontents-hidepagenumbersinweb-property-word.md)|Returns or sets whether page numbers in a table of contents or a table of figures should be hidden when publishing to the Web. Read/write  **Boolean** .|
|[IncludePageNumbers](tableofcontents-includepagenumbers-property-word.md)| **True** if page numbers are included in the table of contents. Read/write **Boolean** .|
|[LowerHeadingLevel](tableofcontents-lowerheadinglevel-property-word.md)|Returns or sets the ending heading level for a table of contents or table of figures. Read/write  **Long** .|
|[Parent](tableofcontents-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **TableOfContents** object.|
|[Range](tableofcontents-range-property-word.md)|Returns a  **Range** object that represents the portion of a document that is contained within the specified table of contents.|
|[RightAlignPageNumbers](tableofcontents-rightalignpagenumbers-property-word.md)| **True** if page numbers are aligned with the right margin in a table of contents. Read/write **Boolean** .|
|[TabLeader](tableofcontents-tableader-property-word.md)|Returns or sets the character between entries and their page numbers in an index, table of authorities, table of contents, or table of figures. Read/write  **[WdTabLeader](wdtableader-enumeration-word.md)** .|
|[TableID](tableofcontents-tableid-property-word.md)|Returns or sets a one-letter identifier that's used to build a table of contents from TOC fields. Read/write  **String** .|
|[UpperHeadingLevel](tableofcontents-upperheadinglevel-property-word.md)|Returns or sets the starting heading level for a table of contents. Read/write  **Long** .|
|[UseFields](tableofcontents-usefields-property-word.md)| **True** if Table of Contents Entry (TC) fields are used to create a table of contents or a table of figures. Read/write **Boolean** .|
|[UseHeadingStyles](tableofcontents-useheadingstyles-property-word.md)| **True** if built-in heading styles are used to create a table of contents. Read/write **Boolean** .|
|[UseHyperlinks](tableofcontents-usehyperlinks-property-word.md)|Returns or sets whether entries in a table of contents should be formatted as hyperlinks when publishing to the Web. Read/write  **Boolean** .|

