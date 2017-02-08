---
title: View Object (Outlook)
keywords: vbaol11.chm2479
f1_keywords:
- vbaol11.chm2479
ms.prod: OUTLOOK
api_name:
- Outlook.View
ms.assetid: 41c8d149-9912-1685-4c8b-3c849cc6f1ed
---


# View Object (Outlook)

Represents a customizable view used to sort, group, and view data.


## Remarks

The  **View** object allows you to create customizable views that allow you to better sort, group and ultimately view data of all different types. There are a variety of different view types that provide the flexibility needed to create and maintain your important data.


- The table view type ( **olTableView** ) allows you to view data in a simple field-based table.
    
- The Calendar view type ( **olCalendarView** ) allows you to view data in a calendar format.
    
- The card view type ( **olCardView** ) allows you to view data in a series of cards. Each card displays the information contained by the item and can be sorted.
    
- The icon view type ( **olIconView** ) allows you to view data as icons, similar to a Windows folder or explorer.
    
- The timeline view type ( **olTimelineView** ) allows you to view data as it is received in a customizable linear time line.
    
Views are defined and customized using the  **View** object's **[XML](http://msdn.microsoft.com/library/view-xml-property-outlook%28Office.15%29.aspx)** property. The **XML** property allows you to create and set a customized XML schema that defines the various features of a view.

Use  **Views** ( _index_ ), where _index_ is the name of the **View** object or its ordinal value, to return a single **View** object.

Use the  **[Add](http://msdn.microsoft.com/library/views-add-method-outlook%28Office.15%29.aspx)** method of the **Views** collection to create a new view.

Always use  **[Save](http://msdn.microsoft.com/library/view-save-method-outlook%28Office.15%29.aspx)** to save a view after you change any property of the view.


## Example

The following example returns a view called Table View and stores it in a variable of type  **View** called objView. Before running this example, make sure a view by the name 'Table View' exists.


```
Sub GetView() 
 
 'Creates a new view 
 
 Dim objName As NameSpace 
 
 Dim objViews As Views 
 
 Dim objView As View 
 
 
 
 Set objName = Application.GetNamespace("MAPI") 
 
 Set objViews = objName.GetDefaultFolder(olFolderInbox).Views 
 
 'Return a view called Table View 
 
 Set objView = objViews.Item("Table View") 
 
End Sub
```

The following example creates a new view of type  **olTableView** called New Table.




```
Sub CreateView() 
 
 'Creates a new view 
 
 Dim objName As NameSpace 
 
 Dim objViews As Views 
 
 Dim objNewView As View 
 
 
 
 Set objName = Application.GetNamespace("MAPI") 
 
 Set objViews = objName.GetDefaultFolder(olFolderInbox).Views 
 
 Set objNewView = objViews.Add(Name:="New Table", _ 
 
 ViewType:=olTableView, SaveOption:=olViewSaveOptionThisFolderEveryone) 
 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Apply](http://msdn.microsoft.com/library/view-apply-method-outlook%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/view-copy-method-outlook%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/view-delete-method-outlook%28Office.15%29.aspx)|
|[GoToDate](http://msdn.microsoft.com/library/view-gotodate-method-outlook%28Office.15%29.aspx)|
|[Reset](http://msdn.microsoft.com/library/view-reset-method-outlook%28Office.15%29.aspx)|
|[Save](http://msdn.microsoft.com/library/view-save-method-outlook%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/view-application-property-outlook%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/view-class-property-outlook%28Office.15%29.aspx)|
|[Filter](http://msdn.microsoft.com/library/view-filter-property-outlook%28Office.15%29.aspx)|
|[Language](http://msdn.microsoft.com/library/view-language-property-outlook%28Office.15%29.aspx)|
|[LockUserChanges](http://msdn.microsoft.com/library/view-lockuserchanges-property-outlook%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/view-name-property-outlook%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/view-parent-property-outlook%28Office.15%29.aspx)|
|[SaveOption](http://msdn.microsoft.com/library/view-saveoption-property-outlook%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/view-session-property-outlook%28Office.15%29.aspx)|
|[Standard](http://msdn.microsoft.com/library/view-standard-property-outlook%28Office.15%29.aspx)|
|[ViewType](http://msdn.microsoft.com/library/view-viewtype-property-outlook%28Office.15%29.aspx)|
|[XML](http://msdn.microsoft.com/library/view-xml-property-outlook%28Office.15%29.aspx)|

## See also


#### Other resources


[View Object Members](http://msdn.microsoft.com/library/view-members-outlook%28Office.15%29.aspx)
[Outlook Object Model Reference](http://msdn.microsoft.com/library/object-model-outlook-vba-reference%28Office.15%29.aspx)
