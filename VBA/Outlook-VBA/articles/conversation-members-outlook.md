---
title: Conversation Members (Outlook)
ms.prod: OUTLOOK
ms.assetid: 09ff1e8e-7c5a-0b1e-e8e2-e259f66f71c8
---


# Conversation Members (Outlook)
Represents a conversation that includes one or more items stored in one or more folders and stores.

Represents a conversation that includes one or more items stored in one or more folders and stores.


## Methods



|**Name**|**Description**|
|:-----|:-----|
|[ClearAlwaysAssignCategories](conversation-clearalwaysassigncategories-method-outlook.md)|Removes all categories from all items in the conversation and stops the action of always assigning categories to items in the conversation.|
|[GetAlwaysAssignCategories](conversation-getalwaysassigncategories-method-outlook.md)|Returns a  **String** that indicates the category or categories that are assigned to all new items that arrive in the conversation.|
|[GetAlwaysDelete](conversation-getalwaysdelete-method-outlook.md)|Returns a constant in the  **[OlAlwaysDeleteConversation](olalwaysdeleteconversation-enumeration-outlook.md)** enumeration that indicates whether all new items that join the conversation are always moved to the **Deleted Items** folder in the specified delivery store.|
|[GetAlwaysMoveToFolder](conversation-getalwaysmovetofolder-method-outlook.md)|Returns a  **[Folder](folder-object-outlook.md)** object that indicates the folder in the specified delivery store to which new items that arrive in the conversation are always moved.|
|[GetChildren](conversation-getchildren-method-outlook.md)|Returns a  **[SimpleItems](simpleitems-object-outlook.md)** collection that contains all items under the specified conversation node.|
|[GetParent](conversation-getparent-method-outlook.md)|Returns the parent item of the specified node in the conversation.|
|[GetRootItems](conversation-getrootitems-method-outlook.md)|Returns a  **[SimpleItems](simpleitems-object-outlook.md)** collection that contains all root items in the conversation.|
|[GetTable](conversation-gettable-method-outlook.md)|Returns a  **[Table](table-object-outlook.md)** object that contains rows that represent all items in the conversation.|
|[MarkAsRead](conversation-markasread-method-outlook.md)|Marks all items in the conversation as read.|
|[MarkAsUnread](conversation-markasunread-method-outlook.md)|Marks all items in the conversation as unread.|
|[SetAlwaysAssignCategories](conversation-setalwaysassigncategories-method-outlook.md)|Applies one or more categories to all existing items and future items of the conversation.|
|[SetAlwaysDelete](conversation-setalwaysdelete-method-outlook.md)|Specifies a setting for the specified delivery store that indicates whether all existing items and all new items that arrive in the conversation are always moved to the Deleted Items folder in the specified delivery store.|
|[SetAlwaysMoveToFolder](conversation-setalwaysmovetofolder-method-outlook.md)|Sets a  **[Folder](folder-object-outlook.md)** object that indicates the folder to which all existing conversation items and new items that arrive in the conversation are always moved.|
|[StopAlwaysDelete](conversation-stopalwaysdelete-method-outlook.md)|Stops the action of always moving conversation items in the specified store to the Deleted Items folder in that store.|
|[StopAlwaysMoveToFolder](conversation-stopalwaysmovetofolder-method-outlook.md)|Stops the action of always moving conversation items in the specified store to a specific folder.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](conversation-application-property-outlook.md)|Returns an  **[Application](application-object-outlook.md)** object that represents the parent Microsoft Outlook application for the **[Conversation](conversation-object-outlook.md)** object. Read-only.|
|[Class](conversation-class-property-outlook.md)|Returns an  **[OlObjectClass](olobjectclass-enumeration-outlook.md)** constant that indicates the object's class. Read-only.|
|[ConversationID](conversation-conversationid-property-outlook.md)|Returns a  **String** that uniquely identifies a **[Conversation](conversation-object-outlook.md)** object. Read-only.|
|[Parent](conversation-parent-property-outlook.md)|Returns the parent  **Object** of the specified **[Conversation](conversation-object-outlook.md)** object. Read-only.|
|[Session](conversation-session-property-outlook.md)|Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.|

