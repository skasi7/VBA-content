---
title: Conversation Object (Outlook)
keywords: vbaol11.chm3388
f1_keywords:
- vbaol11.chm3388
ms.prod: OUTLOOK
api_name:
- Outlook.Conversation
ms.assetid: 2705d38a-ebc0-e5a7-208b-ffe1f5446b1b
---


# Conversation Object (Outlook)

Represents a conversation that includes one or more items stored in one or more folders and stores.


## Remarks

The  **Conversation** object is an abstract, aggregated object. Although a conversation can include items of different types, the **Conversation** object does not correspond to a particular underlying MAPI **IMessage** object.

A conversation represents one or more items in one or more folders and stores. If you move an item in a conversation to the  **Deleted Items** folder and subsequently enumerate the conversation by using the **[GetChildren](http://msdn.microsoft.com/library/conversation-getchildren-method-outlook%28Office.15%29.aspx)**, **[GetRootItems](http://msdn.microsoft.com/library/conversation-getrootitems-method-outlook%28Office.15%29.aspx)**, or **[GetTable](http://msdn.microsoft.com/library/conversation-gettable-method-outlook%28Office.15%29.aspx)** method, the item will not be included in the returned object.

To obtain a  **Conversation** object for an existing conversation, use the **GetConversation** method of the item.

There are actions that you can apply to items in a conversation by calling the  **[SetAlwaysAssignCategories](http://msdn.microsoft.com/library/conversation-setalwaysassigncategories-method-outlook%28Office.15%29.aspx)**, **[SetAlwaysDelete](http://msdn.microsoft.com/library/conversation-setalwaysdelete-method-outlook%28Office.15%29.aspx)**, or **[SetAlwaysMoveToFolder](http://msdn.microsoft.com/library/conversation-setalwaysmovetofolder-method-outlook%28Office.15%29.aspx)** method. Each of these actions is applied to all items in the conversation automatically when the method is called; the action is also applied to future items in the conversation as long as the action is still applicable to the conversation. There is no explicit save method on the **Conversation** object.

Also, when you apply an action to items in a conversation, the corresponding event occurs. For example, the  **[ItemChange](http://msdn.microsoft.com/library/items-itemchange-event-outlook%28Office.15%29.aspx)** event of the **[Items](items-object-outlook.md)** object occurs when you call **SetAlwaysAssignCategories**, and the **[BeforeItemMove](http://msdn.microsoft.com/library/folder-beforeitemmove-event-outlook%28Office.15%29.aspx)** event of the **[Folder](folder-object-outlook.md)** object occurs when you call **SetAlwaysMoveToFolder**.


## Example

The following managed code is written in C#. To run a .NET Framework managed code sample that needs to call into a Component Object Model (COM), you must use an interop assembly that defines and maps managed interfaces to the COM objects in the object model type library. For Outlook, you can use Visual Studio and the Outlook Primary Interop Assembly (PIA). Before you run managed code samples for Outlook 2013, ensure that you have installed the Outlook 2013 PIA and have added a reference to the Microsoft Outlook 15.0 Object Library component in Visual Studio. You should use the following code in the  `ThisAddIn` class of an Outlook add-in (using Office Developer Tools for Visual Studio). The **Application** object in the code must be a trusted Outlook **Application** object provided by `ThisAddIn.Globals`. For more information about using the Outlook PIA to develop managed Outlook solutions, see the  **Welcome to the Outlook Primary Interop Assembly Reference** on MSDN.

The following code example assumes that the selected item in the explorer window is a mail item. The code example gets the conversation that the selected mail item is associated with, and enumerates each item in that conversation, displaying the subject of the item. The  `DemoConversation` method calls the **GetConversation** method of the selected mail item to get the associated **Conversation** object. `DemoConversation` then calls the **[GetTable](http://msdn.microsoft.com/library/conversation-gettable-method-outlook%28Office.15%29.aspx)** and **[GetRootItems](http://msdn.microsoft.com/library/conversation-getrootitems-method-outlook%28Office.15%29.aspx)** methods of the **Conversation** object to get a **[Table](table-object-outlook.md)** object and **[SimpleItems](http://msdn.microsoft.com/library/simpleitems-object-outlook%28Office.15%29.aspx)** collection, respectively. `DemoConversation` calls the recurrent method `EnumerateConversation` to enumerate and display the subject of each item in that conversation.




```C#
void DemoConversation() 
{ 
 object selectedItem = 
 Application.ActiveExplorer().Selection[1]; 
 // This example uses only 
 // MailItem. Other item types such as 
 // MeetingItem and PostItem can participate 
 // in the conversation. 
 if (selectedItem is Outlook.MailItem) 
 { 
 // Cast selectedItem to MailItem. 
 Outlook.MailItem mailItem = 
 selectedItem as Outlook.MailItem; 
 // Determine the store of the mail item. 
 Outlook.Folder folder = mailItem.Parent 
 as Outlook.Folder; 
 Outlook.Store store = folder.Store; 
 if (store.IsConversationEnabled == true) 
 { 
 // Obtain a Conversation object. 
 Outlook.Conversation conv = 
 mailItem.GetConversation(); 
 // Check for null Conversation. 
 if (conv != null) 
 { 
 // Obtain Table that contains rows 
 // for each item in the conversation. 
 Outlook.Table table = conv.GetTable(); 
 Debug.WriteLine("Conversation Items Count: " + 
 table.GetRowCount().ToString()); 
 Debug.WriteLine("Conversation Items from Table:"); 
 while (!table.EndOfTable) 
 { 
 Outlook.Row nextRow = table.GetNextRow(); 
 Debug.WriteLine(nextRow["Subject"] 
 + " Modified: " 
 + nextRow["LastModificationTime"]); 
 } 
 Debug.WriteLine("Conversation Items from Root:"); 
 // Obtain root items and enumerate the conversation. 
 Outlook.SimpleItems simpleItems 
 = conv.GetRootItems(); 
 foreach (object item in simpleItems) 
 { 
 // In this example, only enumerate MailItem type. 
 // Other types such as PostItem or MeetingItem 
 // can appear in the conversation. 
 if (item is Outlook.MailItem) 
 { 
 Outlook.MailItem mail = item 
 as Outlook.MailItem; 
 Outlook.Folder inFolder = 
 mail.Parent as Outlook.Folder; 
 string msg = mail.Subject 
 + " in folder " + inFolder.Name; 
 Debug.WriteLine(msg); 
 } 
 // Call EnumerateConversation 
 // to access child nodes of root items. 
 EnumerateConversation(item, conv); 
 } 
 } 
 } 
 } 
} 
 
 
void EnumerateConversation(object item, 
 Outlook.Conversation conversation) 
{ 
 Outlook.SimpleItems items = 
 conversation.GetChildren(item); 
 if (items.Count > 0) 
 { 
 foreach (object myItem in items) 
 { 
 // In this example, only enumerate MailItem type. 
 // Other types such as PostItem or MeetingItem 
 // can appear in the conversation. 
 if (myItem is Outlook.MailItem) 
 { 
 Outlook.MailItem mailItem = 
 myItem as Outlook.MailItem; 
 Outlook.Folder inFolder = 
 mailItem.Parent as Outlook.Folder; 
 string msg = mailItem.Subject 
 + " in folder " + inFolder.Name; 
 Debug.WriteLine(msg); 
 } 
 // Continue recursion. 
 EnumerateConversation(myItem, conversation); 
 } 
 } 
} 
 

```


## Methods



|**Name**|
|:-----|
|[ClearAlwaysAssignCategories](http://msdn.microsoft.com/library/conversation-clearalwaysassigncategories-method-outlook%28Office.15%29.aspx)|
|[GetAlwaysAssignCategories](http://msdn.microsoft.com/library/conversation-getalwaysassigncategories-method-outlook%28Office.15%29.aspx)|
|[GetAlwaysDelete](http://msdn.microsoft.com/library/conversation-getalwaysdelete-method-outlook%28Office.15%29.aspx)|
|[GetAlwaysMoveToFolder](http://msdn.microsoft.com/library/conversation-getalwaysmovetofolder-method-outlook%28Office.15%29.aspx)|
|[GetChildren](http://msdn.microsoft.com/library/conversation-getchildren-method-outlook%28Office.15%29.aspx)|
|[GetParent](http://msdn.microsoft.com/library/conversation-getparent-method-outlook%28Office.15%29.aspx)|
|[GetRootItems](http://msdn.microsoft.com/library/conversation-getrootitems-method-outlook%28Office.15%29.aspx)|
|[GetTable](http://msdn.microsoft.com/library/conversation-gettable-method-outlook%28Office.15%29.aspx)|
|[MarkAsRead](http://msdn.microsoft.com/library/conversation-markasread-method-outlook%28Office.15%29.aspx)|
|[MarkAsUnread](http://msdn.microsoft.com/library/conversation-markasunread-method-outlook%28Office.15%29.aspx)|
|[SetAlwaysAssignCategories](http://msdn.microsoft.com/library/conversation-setalwaysassigncategories-method-outlook%28Office.15%29.aspx)|
|[SetAlwaysDelete](http://msdn.microsoft.com/library/conversation-setalwaysdelete-method-outlook%28Office.15%29.aspx)|
|[SetAlwaysMoveToFolder](http://msdn.microsoft.com/library/conversation-setalwaysmovetofolder-method-outlook%28Office.15%29.aspx)|
|[StopAlwaysDelete](http://msdn.microsoft.com/library/conversation-stopalwaysdelete-method-outlook%28Office.15%29.aspx)|
|[StopAlwaysMoveToFolder](http://msdn.microsoft.com/library/conversation-stopalwaysmovetofolder-method-outlook%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/conversation-application-property-outlook%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/conversation-class-property-outlook%28Office.15%29.aspx)|
|[ConversationID](http://msdn.microsoft.com/library/conversation-conversationid-property-outlook%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/conversation-parent-property-outlook%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/conversation-session-property-outlook%28Office.15%29.aspx)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/object-model-outlook-vba-reference%28Office.15%29.aspx)
[Conversation Object Members](http://msdn.microsoft.com/library/conversation-members-outlook%28Office.15%29.aspx)
