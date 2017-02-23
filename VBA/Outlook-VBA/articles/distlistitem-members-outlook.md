---
title: DistListItem Members (Outlook)
ms.prod: OUTLOOK
ms.assetid: 3ba4af84-ce84-61d9-1bc9-fab41bf6f125
---


# DistListItem Members (Outlook)
Represents a distribution list in a Contacts folder.

Represents a distribution list in a Contacts folder.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[AfterWrite](distlistitem-afterwrite-event-outlook.md)|Occurs after Microsoft Outlook has saved the item.|
|[AttachmentAdd](distlistitem-attachmentadd-event-outlook.md)|Occurs when an attachment has been added to an instance of the parent object.|
|[AttachmentRead](distlistitem-attachmentread-event-outlook.md)|Occurs when an attachment in an instance of the parent object has been opened for reading.|
|[AttachmentRemove](distlistitem-attachmentremove-event-outlook.md)|Occurs when an attachment has been removed from an instance of the parent object.|
|[BeforeAttachmentAdd](distlistitem-beforeattachmentadd-event-outlook.md)|Occurs before an attachment is added to an instance of the parent object.|
|[BeforeAttachmentPreview](distlistitem-beforeattachmentpreview-event-outlook.md)|Occurs before an attachment associated with an instance of the parent object is previewed.|
|[BeforeAttachmentRead](distlistitem-beforeattachmentread-event-outlook.md)|Occurs before an attachment associated with an instance of the parent object is read from the file system, an attachment stream, or an  **[Attachment](attachment-object-outlook.md)** object.|
|[BeforeAttachmentSave](distlistitem-beforeattachmentsave-event-outlook.md)|Occurs just before an attachment is saved.|
|[BeforeAttachmentWriteToTempFile](distlistitem-beforeattachmentwritetotempfile-event-outlook.md)|Occurs before an attachment associated with an instance of the parent object is written to a temporary file.|
|[BeforeAutoSave](distlistitem-beforeautosave-event-outlook.md)|Occurs before the item is automatically saved by Outlook.|
|[BeforeCheckNames](distlistitem-beforechecknames-event-outlook.md)|Occurs just before Microsoft Outlook starts resolving names in the recipient collection for an item (which is an instance of the parent object).|
|[BeforeDelete](distlistitem-beforedelete-event-outlook.md)|Occurs before an item (which is an instance of the parent object) is deleted.|
|[BeforeRead](distlistitem-beforeread-event-outlook.md)|Occurs before Microsoft Outlook begins to read the properties for the item.|
|[Close](distlistitem-close-event-outlook.md)|Occurs when the inspector associated with an item (which is an instance of the parent object) is being closed.|
|[CustomAction](distlistitem-customaction-event-outlook.md)|Occurs when a custom action of an item (which is an instance of the parent object) executes.|
|[CustomPropertyChange](distlistitem-custompropertychange-event-outlook.md)|Occurs when a custom property of an item (which is an instance of the parent object) is changed. |
|[Forward](distlistitem-forward-event-outlook.md)|Occurs when the user selects the  **Forward** action for an item (which is an instance of the parent object).|
|[Open](distlistitem-open-event-outlook.md)|Occurs when an instance of the parent object is being opened in an  **[Inspector](inspector-object-outlook.md)** .|
|[PropertyChange](distlistitem-propertychange-event-outlook.md)|Occurs when an explicit built-in property (for example,  **[Subject](appointmentitem-subject-property-outlook.md)** ) of an instance of the parent object is changed.|
|[Read](distlistitem-read-event-outlook.md)|Occurs when an instance of the parent object is opened for editing by the user. |
|[ReadComplete](distlistitem-readcomplete-event-outlook.md)|Occurs when Outlook has completed reading the properties of the item.|
|[Reply](distlistitem-reply-event-outlook.md)|Occurs when the user selects the  **Reply** action for an item (which is an instance of the parent object).|
|[ReplyAll](distlistitem-replyall-event-outlook.md)|Occurs when the user selects the  **ReplyAll** action for an item (which is an instance of the parent object).|
|[Send](distlistitem-send-event-outlook.md)|Occurs when the user selects the  **Send** action for an item (which is an instance of the parent object).|
|[Unload](distlistitem-unload-event-outlook.md)|Occurs before an Outlook item is unloaded from memory, either programmatically or by user action. |
|[Write](distlistitem-write-event-outlook.md)|Occurs when an instance of the parent object is saved, either explicitly (for example, using the  **[Save](distlistitem-save-method-outlook.md)** or **[SaveAs](distlistitem-saveas-method-outlook.md)** methods) or implicitly (for example, in response to a prompt when closing the item's inspector).|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[AddMember](distlistitem-addmember-method-outlook.md)|Adds a new member to the specified distribution list. The distribution list contains  **[Recipient](recipient-object-outlook.md)** objects that represent valid e-mail addresses.|
|[AddMembers](distlistitem-addmembers-method-outlook.md)|Adds new members to a distribution list.|
|[ClearTaskFlag](distlistitem-cleartaskflag-method-outlook.md)|Clears the  **[DistListItem](distlistitem-object-outlook.md)** object as a task.|
|[Close](distlistitem-close-method-outlook.md)|Closes and optionally saves changes to the Outlook item.|
|[Copy](distlistitem-copy-method-outlook.md)|Creates another instance of an object.|
|[Delete](distlistitem-delete-method-outlook.md)|Removes the item from the folder that contains the item.|
|[Display](distlistitem-display-method-outlook.md)|Displays a new  **[Inspector](inspector-object-outlook.md)** object for the item.|
|[GetConversation](distlistitem-getconversation-method-outlook.md)|Obtains a  **[Conversation](conversation-object-outlook.md)** object that represents the conversation to which this item belongs.|
|[GetMember](distlistitem-getmember-method-outlook.md)|Returns a  **[Recipient](recipient-object-outlook.md)** object representing a member in a distribution list.|
|[MarkAsTask](distlistitem-markastask-method-outlook.md)|Marks a  **[DistListItem](distlistitem-object-outlook.md)** object as a task and assigns a task interval for the object.|
|[Move](distlistitem-move-method-outlook.md)|Moves a Microsoft Outlook item to a new folder.|
|[PrintOut](distlistitem-printout-method-outlook.md)|Prints the Outlook item using all default settings.The  **PrintOut** method is the only Outlook method that can be used for printing.|
|[RemoveMember](distlistitem-removemember-method-outlook.md)|Removes an individual member from a given distribution list.|
|[RemoveMembers](distlistitem-removemembers-method-outlook.md)|Removes members from a distribution list.|
|[Save](distlistitem-save-method-outlook.md)|Saves the Microsoft Outlook item to the current folder or, if this is a new item, to the Outlook default folder for the item type.|
|[SaveAs](distlistitem-saveas-method-outlook.md)|Saves the Microsoft Outlook item to the specified path and in the format of the specified file type. If the file type is not specified, the MSG format (.msg) is used.|
|[ShowCategoriesDialog](distlistitem-showcategoriesdialog-method-outlook.md)|Displays the  **Show Categories** dialog box, which allows you to select categories that correspond to the subject of the item.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Actions](distlistitem-actions-property-outlook.md)|Returns an  **[Actions](actions-object-outlook.md)** collection that represents all the available actions for the item. Read-only.|
|[Application](distlistitem-application-property-outlook.md)|Returns an  **[Application](application-object-outlook.md)** object that represents the parent Outlook application for the object. Read-only.|
|[Attachments](distlistitem-attachments-property-outlook.md)|Returns an  **[Attachments](attachments-object-outlook.md)** object that represents all the attachments for the specified item. Read-only.|
|[AutoResolvedWinner](distlistitem-autoresolvedwinner-property-outlook.md)|Returns a  **Boolean** that determines if the item is a winner of an automatic conflict resolution. Read-only.|
|[BillingInformation](distlistitem-billinginformation-property-outlook.md)|Returns or sets a  **String** representing the billing information associated with the Outlook item. Read/write.|
|[Body](distlistitem-body-property-outlook.md)|Returns or sets a  **String** representing the clear-text body of the Outlook item. Read/write.|
|[Categories](distlistitem-categories-property-outlook.md)|Returns or sets a  **String** representing the categories assigned to the Outlook item. Read/write.|
|[Class](distlistitem-class-property-outlook.md)|Returns an  **[OlObjectClass](olobjectclass-enumeration-outlook.md)** constant indicating the object's class. Read-only.|
|[Companies](distlistitem-companies-property-outlook.md)|Returns or sets a  **String** representing the names of the companies associated with the Outlook item. Read/write.|
|[Conflicts](distlistitem-conflicts-property-outlook.md)|Return the  **[Conflicts](conflicts-object-outlook.md)** object that represents the items that are in conflict for any Outlook item object. Read-only.|
|[ConversationID](distlistitem-conversationid-property-outlook.md)|Returns a  **String** that uniquely identifies a **[Conversation](conversation-object-outlook.md)** object that the **[DistListItem](distlistitem-object-outlook.md)** object belongs to. Read-only.|
|[ConversationIndex](distlistitem-conversationindex-property-outlook.md)|Returns a  **String** that indicates the relative position of the item within the conversation thread. Read-only.|
|[ConversationTopic](distlistitem-conversationtopic-property-outlook.md)|Returns a  **String** representing the topic of the conversation thread of the Outlook item. Read-only.|
|[CreationTime](distlistitem-creationtime-property-outlook.md)|Returns a  **Date** indicating the creation time for the Outlook item. Read-only.|
|[DLName](distlistitem-dlname-property-outlook.md)|Returns or sets a  **String** representing the display name of a distribution list. Read/write.|
|[DownloadState](distlistitem-downloadstate-property-outlook.md)|Returns a constant that belongs to the  **[OlDownloadState](oldownloadstate-enumeration-outlook.md)** enumeration indicating the download state of the item. Read-only.|
|[EntryID](distlistitem-entryid-property-outlook.md)|Returns a  **String** representing the unique Entry ID of the object. Read-only.|
|[FormDescription](distlistitem-formdescription-property-outlook.md)|Returns the  **[FormDescription](formdescription-object-outlook.md)** object that represents the form description for the specified Outlook item. Read-only.|
|[GetInspector](distlistitem-getinspector-property-outlook.md)|Returns an  **[Inspector](inspector-object-outlook.md)** object that represents an inspector initialized to contain the specified item. Read-only.|
|[Importance](distlistitem-importance-property-outlook.md)|Returns or sets an  **[OlImportance](olimportance-enumeration-outlook.md)** constant indicating the relative importance level for the Outlook item. Read/write.|
|[IsConflict](distlistitem-isconflict-property-outlook.md)|Returns a  **Boolean** that determines if the item is in conflict. Read-only.|
|[IsMarkedAsTask](distlistitem-ismarkedastask-property-outlook.md)|Returns a  **Boolean** value that indicates whether the **[DistListItem](distlistitem-object-outlook.md)** is marked as a task. Read-only.|
|[ItemProperties](distlistitem-itemproperties-property-outlook.md)|Returns an  **[ItemProperties](itemproperties-object-outlook.md)** collection that represents all standard and user-defined properties associated with the Outlook item. Read-only.|
|[LastModificationTime](distlistitem-lastmodificationtime-property-outlook.md)|Returns a  **Date** specifying the date and time that the Outlook item was last modified. Read-only.|
|[MarkForDownload](distlistitem-markfordownload-property-outlook.md)|Returns or sets an  **[OlRemoteStatus](olremotestatus-enumeration-outlook.md)** constant that determines the status of an item once it is received by a remote user. Read/write.|
|[MemberCount](distlistitem-membercount-property-outlook.md)|Returns a  **Long** indicating the number of members in a distribution list. Read-only.|
|[MessageClass](distlistitem-messageclass-property-outlook.md)|Returns or sets a  **String** representing the message class for the Outlook item. Read/write.|
|[Mileage](distlistitem-mileage-property-outlook.md)|Returns or sets a  **String** representing the mileage for an item. Read/write.|
|[NoAging](distlistitem-noaging-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** to not age the Outlook item. Read/write.|
|[OutlookInternalVersion](distlistitem-outlookinternalversion-property-outlook.md)|Returns a  **Long** representing the build number of the Outlook application for an Outlook item. Read-only.|
|[OutlookVersion](distlistitem-outlookversion-property-outlook.md)|Returns a  **String** indicating the major and minor version number of the Outlook application for an Outlook item. Read-only.|
|[Parent](distlistitem-parent-property-outlook.md)|Returns the parent  **Object** of the specified object. Read-only.|
|[PropertyAccessor](distlistitem-propertyaccessor-property-outlook.md)|Returns a  **[PropertyAccessor](propertyaccessor-object-outlook.md)** object that supports creating, getting, setting, and deleting properties of the parent **[DistListItem](distlistitem-object-outlook.md)** object. Read-only.|
|[ReminderOverrideDefault](distlistitem-reminderoverridedefault-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** if the reminder overrides the default reminder behavior for the item. Read/write.|
|[ReminderPlaySound](distlistitem-reminderplaysound-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** if the reminder should play a sound when it occurs for this item. Read/write.|
|[ReminderSet](distlistitem-reminderset-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** if a reminder has been set for this item. Read/write.|
|[ReminderSoundFile](distlistitem-remindersoundfile-property-outlook.md)|Returns or sets a  **String** indicating the path and file name of the sound file to play when the reminder occurs for the Outlook item. Read/write.|
|[ReminderTime](distlistitem-remindertime-property-outlook.md)|Returns or sets a  **Date** indicating the date and time at which the reminder should occur for the specified item. Read/write.|
|[RTFBody](distlistitem-rtfbody-property-outlook.md)|Returns or sets a  **Byte** array that represents the body of the Microsoft Outlook item in Rich Text Format. Read/write.|
|[Saved](distlistitem-saved-property-outlook.md)|Returns a  **Boolean** value that is **True** if the Outlook item has not been modified since the last save. Read-only.|
|[Sensitivity](distlistitem-sensitivity-property-outlook.md)|Returns or sets a constant in the  **[OlSensitivity](olsensitivity-enumeration-outlook.md)** enumeration indicating the sensitivity for the Outlook item. Read/write.|
|[Session](distlistitem-session-property-outlook.md)|Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.|
|[Size](distlistitem-size-property-outlook.md)|Returns a  **Long** indicating the size (in bytes) of the Outlook item. Read-only.|
|[Subject](distlistitem-subject-property-outlook.md)|Returns or sets a  **String** indicating the subject for the Outlook item. Read/write.|
|[TaskCompletedDate](distlistitem-taskcompleteddate-property-outlook.md)|Returns or sets a  **Date** value that represents the completion date of the task for this **[DistListItem](distlistitem-object-outlook.md)** . Read/write.|
|[TaskDueDate](distlistitem-taskduedate-property-outlook.md)|Returns or sets a  **Date** value that represents the due date of the task for this **[DistListItem](distlistitem-object-outlook.md)** . Read/write.|
|[TaskStartDate](distlistitem-taskstartdate-property-outlook.md)|Returns or sets a  **Date** value that represents the start date of the task for this **[DistListItem](distlistitem-object-outlook.md)** object. Read/write.|
|[TaskSubject](distlistitem-tasksubject-property-outlook.md)|Returns or sets a  **String** value that represents the subject of the task for the **[DistListItem](distlistitem-object-outlook.md)** object. Read/write.|
|[ToDoTaskOrdinal](distlistitem-todotaskordinal-property-outlook.md)|Returns or sets a  **Date** value that represents the ordinal value of the task for the **[DistListItem](distlistitem-object-outlook.md)** . Read/write.|
|[UnRead](distlistitem-unread-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** if the Outlook item has not been opened (read). Read/write.|
|[UserProperties](distlistitem-userproperties-property-outlook.md)|Returns the  **[UserProperties](userproperties-object-outlook.md)** collection that represents all the user properties for the Outlook item. Read-only.|

