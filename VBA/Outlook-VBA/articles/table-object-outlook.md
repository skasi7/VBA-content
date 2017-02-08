---
title: Table Object (Outlook)
keywords: vbaol11.chm3166
f1_keywords:
- vbaol11.chm3166
ms.prod: OUTLOOK
api_name:
- Outlook.Table
ms.assetid: 0affaafd-93fe-227a-acee-e09a86cadc20
---


# Table Object (Outlook)

Represents a set of item data from a  **[Folder](folder-object-outlook.md)** or **[Search](search-object-outlook.md)** object, with items as rows of the table and properties as columns of the table.


## Remarks

The  **Table** represents a read-only dynamic rowset of data in a **Folder** or **Search** object. You can use **[Folder.GetTable](http://msdn.microsoft.com/library/folder-gettable-method-outlook%28Office.15%29.aspx)** or **[Search.GetTable](http://msdn.microsoft.com/library/search-gettable-method-outlook%28Office.15%29.aspx)** to obtain a **Table** object that represents a set of items in a folder or search folder. If the **Table** object is obtained from **Folder.GetTable**, you can further specify a filter (in **[Table.Restrict](http://msdn.microsoft.com/library/table-restrict-method-outlook%28Office.15%29.aspx)** ) to obtain a subset of the items in the folder. If you do not specify any filter, you will obtain all the items in the folder.

By default, each item in the returned  **Table** contains only a default subset of its properties. You can regard each row of a **Table** as an item in the folder, each column as a property of the item, and the **Table** as an in-memory lightweight rowset that allows fast enumeration and filtering of items in the folder. Although additions and deletions of the underlying folder are reflected by the rows in the **Table**, the **Table** does not support any events for adding, changing, and removing of rows. If you require a writeable object from the **Table** row, obtain the Entry ID for that row from the default EntryID column in the **Table** and then use the **[GetItemFromID](http://msdn.microsoft.com/library/namespace-getitemfromid-method-outlook%28Office.15%29.aspx)** method of the **[NameSpace](namespace-object-outlook.md)** object to obtain a full item, such as a **[MailItem](http://msdn.microsoft.com/library/mailitem-object-outlook%28Office.15%29.aspx)** or **[ContactItem](contactitem-object-outlook.md)**, that supports read-write operations. For more information on default columns in a **Table**, see[Default Properties Displayed in a Table Object](http://msdn.microsoft.com/library/default-properties-displayed-in-a-table-object%28Office.15%29.aspx).

 For more information on the **Table** object, see[Enumerating, Searching, and Filtering Items in a Folder](http://msdn.microsoft.com/library/enumerating-searching-and-filtering-items-in-a-folder%28Office.15%29.aspx).


## Example

The following code sample illustrates how the  **Table** object can return a filtered set of items based on their **LastModificationTime** property. It also shows how to list the default properties as well as specific properties of the items.


```
Sub DemoTable() 
 
 'Declarations 
 
 Dim Filter As String 
 
 Dim oRow As Outlook.Row 
 
 Dim oTable As Outlook.Table 
 
 Dim oFolder As Outlook.Folder 
 
 
 
 'Get a Folder object for the Inbox 
 
 Set oFolder = Application.Session.GetDefaultFolder(olFolderInbox) 
 
 
 
 'Define Filter to obtain items last modified after May 1, 2005 
 
 Filter = "[LastModificationTime] > '5/1/2005'" 
 
 'Restrict with Filter 
 
 Set oTable = oFolder.GetTable(Filter) 
 
 
 
 'Remove all columns in the default column set 
 
 oTable.Columns.RemoveAll 
 
 'Specify desired properties 
 
 With oTable.Columns 
 
 .Add ("Subject") 
 
 .Add ("LastModificationTime") 
 
 'PR_ATTR_HIDDEN referenced by the MAPI proptag namespace 
 
 .Add ("http://schemas.microsoft.com/mapi/proptag/0x10F4000B") 
 
 End With 
 
 
 
 'Enumerate the table using test for EndOfTable 
 
 Do Until (oTable.EndOfTable) 
 
 Set oRow = oTable.GetNextRow() 
 
 Debug.Print (oRow("Subject")) 
 
 Debug.Print (oRow("LastModificationTime")) 
 
 Debug.Print (oRow("http://schemas.microsoft.com/mapi/proptag/0x10F4000B")) 
 
 Loop 
 
End Sub
```


## Methods



|**Name**|
|:-----|
|[FindNextRow](http://msdn.microsoft.com/library/table-findnextrow-method-outlook%28Office.15%29.aspx)|
|[FindRow](http://msdn.microsoft.com/library/table-findrow-method-outlook%28Office.15%29.aspx)|
|[GetArray](http://msdn.microsoft.com/library/table-getarray-method-outlook%28Office.15%29.aspx)|
|[GetNextRow](http://msdn.microsoft.com/library/table-getnextrow-method-outlook%28Office.15%29.aspx)|
|[GetRowCount](http://msdn.microsoft.com/library/table-getrowcount-method-outlook%28Office.15%29.aspx)|
|[MoveToStart](http://msdn.microsoft.com/library/table-movetostart-method-outlook%28Office.15%29.aspx)|
|[Restrict](http://msdn.microsoft.com/library/table-restrict-method-outlook%28Office.15%29.aspx)|
|[Sort](http://msdn.microsoft.com/library/table-sort-method-outlook%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/table-application-property-outlook%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/table-class-property-outlook%28Office.15%29.aspx)|
|[Columns](http://msdn.microsoft.com/library/table-columns-property-outlook%28Office.15%29.aspx)|
|[EndOfTable](http://msdn.microsoft.com/library/table-endoftable-property-outlook%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/table-parent-property-outlook%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/table-session-property-outlook%28Office.15%29.aspx)|

## See also


#### Other resources


[Table Object Members](http://msdn.microsoft.com/library/table-members-outlook%28Office.15%29.aspx)
[Outlook Object Model Reference](http://msdn.microsoft.com/library/object-model-outlook-vba-reference%28Office.15%29.aspx)
