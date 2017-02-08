---
title: Table Members (Outlook)
ms.prod: OUTLOOK
ms.assetid: bd9db35d-0738-22cf-a936-425d5a0ead87
---


# Table Members (Outlook)
Represents a set of item data from a  **[Folder](folder-object-outlook.md)** or **[Search](search-object-outlook.md)** object, with items as rows of the table and properties as columns of the table.

Represents a set of item data from a  **[Folder](folder-object-outlook.md)** or **[Search](search-object-outlook.md)** object, with items as rows of the table and properties as columns of the table.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[FindNextRow](table-findnextrow-method-outlook.md)|Finds the next row in the  **[Table](table-object-outlook.md)** that meets the criteria specified in a preceding **[Table.FindRow](table-findrow-method-outlook.md)** .|
|[FindRow](table-findrow-method-outlook.md)|Finds the first row in the  **[Table](table-object-outlook.md)** that meets the criteria specified in _Filter_ .|
|[GetArray](table-getarray-method-outlook.md)|Obtains a two-dimensional array that contains a set of row and column values from the  **[Table](table-object-outlook.md)** .|
|[GetNextRow](table-getnextrow-method-outlook.md)|Moves the current row to the next row in the  **[Table](table-object-outlook.md)** and obtains that row in the **Table** .|
|[GetRowCount](table-getrowcount-method-outlook.md)|Obtains the number of rows in the  **[Table](table-object-outlook.md)** .|
|[MoveToStart](table-movetostart-method-outlook.md)|Moves the current row of the  **[Table](table-object-outlook.md)** to just before the first row of the **Table** .|
|[Restrict](table-restrict-method-outlook.md)|Applies a filter to the rows in the  **[Table](table-object-outlook.md)** and obtains a new **Table** object.|
|[Sort](table-sort-method-outlook.md)|Sorts the rows of the  **[Table](table-object-outlook.md)** by the property specified in _SortProperty_ and resets the current row to just before the first row in the **Table** .|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](table-application-property-outlook.md)|Returns an  **[Application](application-object-outlook.md)** object that represents the parent application (Outlook) for the **[Table](table-object-outlook.md)** object. Read-only.|
|[Class](table-class-property-outlook.md)|Returns a constant in the  **[OlObjectClass](olobjectclass-enumeration-outlook.md)** enumeration indicating the class of the **[Table](table-object-outlook.md)** object. Read-only.|
|[Columns](table-columns-property-outlook.md)|Returns a  **[Columns](columns-object-outlook.md)** collection object that contains the columns defined for the **[Table](table-object-outlook.md)** . Read-only.|
|[EndOfTable](table-endoftable-property-outlook.md)|Returns a  **Boolean** that indicates whether the current row is positioned after the last row in the **[Table](table-object-outlook.md)** object. Read-only.|
|[Parent](table-parent-property-outlook.md)|Returns the parent  **Object** of the **[Table](table-object-outlook.md)** object. Read-only.|
|[Session](table-session-property-outlook.md)|Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.|

