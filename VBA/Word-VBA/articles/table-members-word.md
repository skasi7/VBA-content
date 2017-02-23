---
title: Table Members (Word)
ms.prod: WORD
ms.assetid: 5367ee92-b5a3-92c7-787b-46a302586a0d
---


# Table Members (Word)
Represents a single table. The  **Table** object is a member of the **[Tables](tables-object-word.md)** collection. The **Tables** collection includes all the tables in the specified selection, range, or document.

Represents a single table. The  **Table** object is a member of the **[Tables](tables-object-word.md)** collection. The **Tables** collection includes all the tables in the specified selection, range, or document.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[ApplyStyleDirectFormatting](table-applystyledirectformatting-method-word.md)|Applies the specified style but maintains any formatting that a user directly applies.|
|[AutoFitBehavior](table-autofitbehavior-method-word.md)|Determines how Microsoft Word resizes a table when the AutoFit feature is used.|
|[AutoFormat](table-autoformat-method-word.md)|Applies a predefined look to a table.|
|[Cell](table-cell-method-word.md)|Returns a  **Cell** object that represents a cell in a table.|
|[ConvertToText](table-converttotext-method-word.md)|Converts a table to text and returns a  **Range** object that represents the delimited text.|
|[Delete](table-delete-method-word.md)|Deletes the specified table.|
|[Select](table-select-method-word.md)|Selects the specified table.|
|[Sort](table-sort-method-word.md)|Sorts the specified table.|
|[SortAscending](table-sortascending-method-word.md)|Sorts paragraphs or table rows in ascending alphanumeric order.|
|[SortDescending](table-sortdescending-method-word.md)|Sorts table rows in descending alphanumeric order.|
|[Split](table-split-method-word.md)|Inserts an empty paragraph immediately above the specified row in the table, and returns a  **Table** object that contains both the specified row and the rows that follow it.|
|[UpdateAutoFormat](table-updateautoformat-method-word.md)|Updates the table with the characteristics of a predefined table format.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[AllowAutoFit](table-allowautofit-property-word.md)|Allows Microsoft Word to automatically resize cells in a table to fit their contents. Read/write  **Boolean** .|
|[Application](table-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[ApplyStyleColumnBands](table-applystylecolumnbands-property-word.md)|Returns or sets a  **Boolean** that represents whether to apply style bands to the columns in a table if an applied preset table style provides style banding for columns. Read/write.|
|[ApplyStyleFirstColumn](table-applystylefirstcolumn-property-word.md)| **True** for Microsoft Word to apply first-column formatting to the first column of the specified table. Read/write **Boolean** .|
|[ApplyStyleHeadingRows](table-applystyleheadingrows-property-word.md)| **True** for Microsoft Word to apply heading-row formatting to the first row of the selected table. Read/write **Boolean** .|
|[ApplyStyleLastColumn](table-applystylelastcolumn-property-word.md)| **True** for Microsoft Word to apply last-column formatting to the last column of the specified table. Read/write **Boolean** .|
|[ApplyStyleLastRow](table-applystylelastrow-property-word.md)| **True** for Microsoft Word to apply last-row formatting to the last row of the specified table. Read/write **Boolean** .|
|[ApplyStyleRowBands](table-applystylerowbands-property-word.md)|Returns or sets a  **Boolean** that represents whether to apply style bands to the rows in a table if an applied preset table style provides style banding for rows. Read/write.|
|[AutoFormatType](table-autoformattype-property-word.md)|Returns the type of automatic formatting that's been applied to the specified table. Read-only  **Long** .|
|[Borders](table-borders-property-word.md)|Returns a  **[Borders](borders-object-word.md)** collection that represents all the borders for the specified object.|
|[BottomPadding](table-bottompadding-property-word.md)|Returns or sets the amount of space (in points) to add below the contents of a single cell or all the cells in a table. Read/write  **Single** .|
|[Columns](table-columns-property-word.md)|Returns a  **[Columns](columns-object-word.md)** collection that represents all the table columns in the table. Read-only.|
|[Creator](table-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[Descr](table-descr-property-word.md)|Returns or sets a  **String** that contains a description for the specified table. Read/write.|
|[ID](table-id-property-word.md)|Returns or sets the identifying label for the specified table when the document is saved as a Web page. Read/write  **String** .|
|[LeftPadding](table-leftpadding-property-word.md)|Returns or sets the amount of space (in points) to add to the left of the contents of all the cells in a table. Read/write  **Single** .|
|[NestingLevel](table-nestinglevel-property-word.md)|Returns the nesting level of the specified table. Read-only  **Long** .|
|[Parent](table-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **Table** object.|
|[PreferredWidth](table-preferredwidth-property-word.md)|Returns or sets the preferred width (in points or as a percentage of the window width) for the specified table. Read/write  **Single** .|
|[PreferredWidthType](table-preferredwidthtype-property-word.md)|Returns or sets the preferred unit of measurement to use for the width of the specified table. Read/write  **[WdPreferredWidthType](wdpreferredwidthtype-enumeration-word.md)** .|
|[Range](table-range-property-word.md)|Returns a  **Range** object that represents the portion of a document that is contained within the specified table.|
|[RightPadding](table-rightpadding-property-word.md)|Returns or sets the amount of space (in points) to add to the right of the contents of all the cells in a table. Read/write  **Single** .|
|[Rows](table-rows-property-word.md)|Returns a  **Rows** collection that represents all the table rows within a table. Read-only.|
|[Shading](table-shading-property-word.md)|Returns a  **Shading** object that refers to the shading formatting for the specified object.|
|[Spacing](table-spacing-property-word.md)|Returns or sets the spacing (in points) between the cells in a table. Read/write  **Single** .|
|[Style](table-style-property-word.md)|Returns or sets the style for the specified table. Read/write  **Variant** .|
|[TableDirection](table-tabledirection-property-word.md)|Returns or sets the direction in which Microsoft Word orders cells in the specified table. Read/write  **[WdTableDirection](wdtabledirection-enumeration-word.md)** .|
|[Tables](table-tables-property-word.md)|Returns a  **[Tables](tables-object-word.md)** collection that represents all the tables nested within the specified table. Read-only.|
|[Title](table-title-property-word.md)|Returns or sets a  **String** that contains a title for the specified table. Read/write.|
|[TopPadding](table-toppadding-property-word.md)|Returns or sets the amount of space (in points) to add above the contents of all the cells in a table. Read/write  **Single** .|
|[Uniform](table-uniform-property-word.md)| **True** if all the rows in a table have the same number of columns. Read-only **Boolean** .|

