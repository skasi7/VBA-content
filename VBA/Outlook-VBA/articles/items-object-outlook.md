---
title: Items Object (Outlook)
keywords: vbaol11.chm2998
f1_keywords:
- vbaol11.chm2998
ms.prod: OUTLOOK
api_name:
- Outlook.Items
ms.assetid: 3a99730b-e62a-5ca6-f6ec-911c95173242
---


# Items Object (Outlook)

Contains a collection of [Outlook item objects](http://msdn.microsoft.com/library/outlook-item-objects%28Office.15%29.aspx) in a folder.


## Remarks

Use the  **[Items](http://msdn.microsoft.com/library/folder-items-property-outlook%28Office.15%29.aspx)** property to return the **Items** object of a **[Folder](folder-object-outlook.md)** object.

Use  **Items** ( _index_ ), where _index_ is the name or index number, to return a single Outlook item.


 **Note**  The index for the  **Items** collection starts at 1, and the items in the **Items** collection object are not guaranteed to be in any particular order.


## Example

The following Microsoft Visual Basic for Applications (VBA) example returns the first item in the  **Inbox** with the Subject "Need your advice."






```
Sub GetItem() 
 
 Dim myNameSpace As Outlook.NameSpace 
 
 Dim myFolder As Outlook.Folder 
 
 Dim myItem As Object 
 
 
 
 Set myNameSpace = Application.GetNameSpace("MAPI") 
 
 Set myFolder = _ 
 
 myNameSpace.GetDefaultFolder(olFolderInbox) 
 
 Set myItem = myFolder.Items("Need your advice") 
 
 myItem.Display 
 
End sub
```

The following VBA example returns the first item in the  **Inbox**. In Microsoft Office Outlook 2003 or later, the  **Items** object returns the items in an Offline Folders file (.ost) in the reverse order.






```
Sub GetItem() 
 
 Dim myNameSpace As Outlook.NameSpace 
 
 Dim myFolder As Outlook.Folder 
 
 Dim myItem As Object 
 
 
 
 Set myNameSpace = Application.GetNameSpace("MAPI") 
 
 Set myFolder = _ 
 
 myNameSpace.GetDefaultFolder(olFolderInbox) 
 
 Set myItem = myFolder.Items(1) 
 
 myItem.Display 
 
End sub
```


## Events



|**Name**|
|:-----|
|[ItemAdd](http://msdn.microsoft.com/library/items-itemadd-event-outlook%28Office.15%29.aspx)|
|[ItemChange](http://msdn.microsoft.com/library/items-itemchange-event-outlook%28Office.15%29.aspx)|
|[ItemRemove](http://msdn.microsoft.com/library/items-itemremove-event-outlook%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/items-add-method-outlook%28Office.15%29.aspx)|
|[Find](http://msdn.microsoft.com/library/items-find-method-outlook%28Office.15%29.aspx)|
|[FindNext](http://msdn.microsoft.com/library/items-findnext-method-outlook%28Office.15%29.aspx)|
|[GetFirst](http://msdn.microsoft.com/library/items-getfirst-method-outlook%28Office.15%29.aspx)|
|[GetLast](http://msdn.microsoft.com/library/items-getlast-method-outlook%28Office.15%29.aspx)|
|[GetNext](http://msdn.microsoft.com/library/items-getnext-method-outlook%28Office.15%29.aspx)|
|[GetPrevious](http://msdn.microsoft.com/library/items-getprevious-method-outlook%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/items-item-method-outlook%28Office.15%29.aspx)|
|[Remove](http://msdn.microsoft.com/library/items-remove-method-outlook%28Office.15%29.aspx)|
|[ResetColumns](http://msdn.microsoft.com/library/items-resetcolumns-method-outlook%28Office.15%29.aspx)|
|[Restrict](http://msdn.microsoft.com/library/items-restrict-method-outlook%28Office.15%29.aspx)|
|[SetColumns](http://msdn.microsoft.com/library/items-setcolumns-method-outlook%28Office.15%29.aspx)|
|[Sort](http://msdn.microsoft.com/library/items-sort-method-outlook%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/items-application-property-outlook%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/items-class-property-outlook%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/items-count-property-outlook%28Office.15%29.aspx)|
|[IncludeRecurrences](http://msdn.microsoft.com/library/items-includerecurrences-property-outlook%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/items-parent-property-outlook%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/items-session-property-outlook%28Office.15%29.aspx)|

## See also


#### Other resources


[Items Object Members](http://msdn.microsoft.com/library/items-members-outlook%28Office.15%29.aspx)
[Outlook Object Model Reference](http://msdn.microsoft.com/library/object-model-outlook-vba-reference%28Office.15%29.aspx)
