---
title: Index Members (Word)
ms.prod: WORD
ms.assetid: de9f0a3c-dd30-84bd-e122-2d20fa6b3d37
---


# Index Members (Word)
Represents a single index. The  **Index** object is a member of the **Indexes** collection. The **[Indexes](indexes-object-word.md)** collection includes all the indexes in the specified document.

Represents a single index. The  **Index** object is a member of the **Indexes** collection. The **[Indexes](indexes-object-word.md)** collection includes all the indexes in the specified document.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Delete](index-delete-method-word.md)|Deletes the specified index.|
|[Update](index-update-method-word.md)|Updates the entries shown in specified index.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[AccentedLetters](index-accentedletters-property-word.md)| **True** if the specified index contains separate headings for accented letters (for example, words that begin with "À" are under one heading and words that begin with "A" are under another). Read/write **Boolean** .|
|[Application](index-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[Creator](index-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[Filter](index-filter-property-word.md)|Returns or sets a value that specifies how Microsoft Word classifies the first character of entries in the specified index.read/write  **Long** . Can be one of the following **wdIndexFilter** constants.|
|[HeadingSeparator](index-headingseparator-property-word.md)|Returns or sets the text between alphabetical groups (entries that start with the same letter) in the index. Corresponds to the \h switch for an INDEX field. Read/write  **WdHeadingSeparator** .|
|[IndexLanguage](index-indexlanguage-property-word.md)|Returns or sets a  **[WdLanguageID](wdlanguageid-enumeration-word.md)** constant that represents the sorting language to use for the specified index. Read/write .|
|[NumberOfColumns](index-numberofcolumns-property-word.md)|Sets or returns the number of columns for each page of an index. Read/write  **Long** .|
|[Parent](index-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **Index** object.|
|[Range](index-range-property-word.md)|Returns a  **Range** object that represents the portion of a document that is contained within an index.|
|[RightAlignPageNumbers](index-rightalignpagenumbers-property-word.md)| **True** if page numbers are aligned with the right margin in an index. Read/write **Boolean** .|
|[SortBy](index-sortby-property-word.md)|Returns or sets the sorting criteria for the specified index. Read/write  **WdIndexSortBy** .|
|[TabLeader](index-tableader-property-word.md)|Returns or sets the leader character between entries in an index and their associated page numbers. Read/write  **WdTabLeader** .|
|[Type](index-type-property-word.md)|Returns or sets the index type. Read/write  **[WdIndexType](wdindextype-enumeration-word.md)** .|

