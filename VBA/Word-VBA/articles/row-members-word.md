---
title: Row Members (Word)
ms.prod: WORD
ms.assetid: 3ac6ec58-8e33-7e98-33b6-861a7aa7e80f
---


# Row Members (Word)
Represents a row in a table. The  **Row** object is a member of the **[Rows](rows-object-word.md)** collection. The **Rows** collection includes all the rows in the specified selection, range, or table.

Represents a row in a table. The  **Row** object is a member of the **[Rows](rows-object-word.md)** collection. The **Rows** collection includes all the rows in the specified selection, range, or table.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[ConvertToText](row-converttotext-method-word.md)|Converts a table to text and returns a  **Range** object that represents the delimited text.|
|[Delete](row-delete-method-word.md)|Deletes the specified table row.|
|[Select](row-select-method-word.md)|Selects the specified table row.|
|[SetHeight](row-setheight-method-word.md)|Sets the height of a table row.|
|[SetLeftIndent](row-setleftindent-method-word.md)|Sets the indentation for a row in a table.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Alignment](row-alignment-property-word.md)|Returns or sets a  **WdRowAlignment** constant that represents the alignment for the specified rows. Read/write.|
|[AllowBreakAcrossPages](row-allowbreakacrosspages-property-word.md)| **True** if the text in a table row or rows are allowed to split across a page break. Read/write **Long** .|
|[Application](row-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[Borders](row-borders-property-word.md)|Returns a  **[Borders](borders-object-word.md)** collection that represents all the borders for the specified object.|
|[Cells](row-cells-property-word.md)|Returns a  **[Cells](cells-object-word.md)** collection that represents the table cells in a column, row, selection, or range. Read-only.|
|[Creator](row-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[HeadingFormat](row-headingformat-property-word.md)| **True** if the specified row or rows are formatted as a table heading. Rows formatted as table headings are repeated when a table spans more than one page. Can be **True** , **False** or **wdUndefined** . Read/write **Long** .|
|[Height](row-height-property-word.md)|Returns or sets the height of the specified row in a table. Read/write Single.|
|[HeightRule](row-heightrule-property-word.md)|Returns or sets the rule for determining the height of the specified cells or rows. Read/write  **WdRowHeightRule** .|
|[ID](row-id-property-word.md)|Returns or sets the identifying label for the specified table row when the document is saved as a Web page. Read/write  **String** .|
|[Index](row-index-property-word.md)|Returns a  **Long** that represents the position of an item in a collection. Read-only.|
|[IsFirst](row-isfirst-property-word.md)| **True** if the specified row is the first one in the table. Read-only **Boolean** .|
|[IsLast](row-islast-property-word.md)| **True** if the specified row is the last one in the table. Read-only **Boolean** .|
|[LeftIndent](row-leftindent-property-word.md)|Returns or sets a  **Single** that represents the left indent value (in points) for the specified table row. Read/write.|
|[NestingLevel](row-nestinglevel-property-word.md)|Returns the nesting level of the specified table row. Read-only  **Long** .|
|[Next](row-next-property-word.md)|Returns a  **Row** object that represents the table row that is next in the collection of rows in a table. Read-only.|
|[Parent](row-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **Row** object.|
|[Previous](row-previous-property-word.md)|Returns a  **Row** object that represents the table row that is previous to the specified row. Read-only.|
|[Range](row-range-property-word.md)|Returns a  **Range** object that represents the portion of a document that is contained within the specified table row.|
|[Shading](row-shading-property-word.md)|Returns a  **[Shading](shading-object-word.md)** object that refers to the shading formatting for the specified object.|
|[SpaceBetweenColumns](row-spacebetweencolumns-property-word.md)|Returns or sets the distance (in points) between text in adjacent columns of the specified row or rows. Read/write  **Single** .|

