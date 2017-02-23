---
title: PostItem Members (Outlook)
ms.prod: OUTLOOK
ms.assetid: 5b150db1-c96d-0721-ec36-d5b5ebc20fd8
---


# PostItem Members (Outlook)
Represents a post in a public folder that others may browse.

Represents a post in a public folder that others may browse.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[AfterWrite](postitem-afterwrite-event-outlook.md)|Occurs after Microsoft Outlook has saved the item.|
|[AttachmentAdd](postitem-attachmentadd-event-outlook.md)|Occurs when an attachment has been added to an instance of the parent object.|
|[AttachmentRead](postitem-attachmentread-event-outlook.md)|Occurs when an attachment in an instance of the parent object has been opened for reading.|
|[AttachmentRemove](postitem-attachmentremove-event-outlook.md)|Occurs when an attachment has been removed from an instance of the parent object.|
|[BeforeAttachmentAdd](postitem-beforeattachmentadd-event-outlook.md)|Occurs before an attachment is added to an instance of the parent object.|
|[BeforeAttachmentPreview](postitem-beforeattachmentpreview-event-outlook.md)|Occurs before an attachment associated with an instance of the parent object is previewed.|
|[BeforeAttachmentRead](postitem-beforeattachmentread-event-outlook.md)|Occurs before an attachment associated with an instance of the parent object is read from the file system, an attachment stream, or an  **[Attachment](attachment-object-outlook.md)** object.|
|[BeforeAttachmentSave](postitem-beforeattachmentsave-event-outlook.md)|Occurs just before an attachment is saved.|
|[BeforeAttachmentWriteToTempFile](postitem-beforeattachmentwritetotempfile-event-outlook.md)|Occurs before an attachment associated with an instance of the parent object is written to a temporary file.|
|[BeforeAutoSave](postitem-beforeautosave-event-outlook.md)|Occurs before the item is automatically saved by Outlook.|
|[BeforeCheckNames](postitem-beforechecknames-event-outlook.md)|Occurs just before Microsoft Outlook starts resolving names in the recipient collection for an item (which is an instance of the parent object).|
|[BeforeDelete](postitem-beforedelete-event-outlook.md)|Occurs before an item (which is an instance of the parent object) is deleted.|
|[BeforeRead](postitem-beforeread-event-outlook.md)|Occurs before Microsoft Outlook begins to read the properties for the item.|
|[Close](postitem-close-event-outlook.md)|Occurs when the inspector associated with an item (which is an instance of the parent object) is being closed.|
|[CustomAction](postitem-customaction-event-outlook.md)|Occurs when a custom action of an item (which is an instance of the parent object) executes.|
|[CustomPropertyChange](postitem-custompropertychange-event-outlook.md)|Occurs when a custom property of an item (which is an instance of the parent object) is changed. |
|[Forward](postitem-forward-event-outlook.md)|Occurs when the user selects the  **Forward** action for an item, or when the **Forward** method is called for the item, which is an instance of the parent object.|
|[Open](postitem-open-event-outlook.md)|Occurs when an instance of the parent object is being opened in an  **[Inspector](inspector-object-outlook.md)** .|
|[PropertyChange](postitem-propertychange-event-outlook.md)|Occurs when an explicit built-in property (for example,  **[Subject](appointmentitem-subject-property-outlook.md)** ) of an instance of the parent object is changed.|
|[Read](postitem-read-event-outlook.md)|Occurs when an instance of the parent object is opened for editing by the user. |
|[ReadComplete](postitem-readcomplete-event-outlook.md)|Occurs when Outlook has completed reading the properties of the item.|
|[Reply](postitem-reply-event-outlook.md)|Occurs when the user selects the  **Reply** action for an item, or when the **Reply** method is called for the item, which is an instance of the parent object.|
|[ReplyAll](postitem-replyall-event-outlook.md)|Occurs when the user selects the  **ReplyAll** action for an item (which is an instance of the parent object).|
|[Send](postitem-send-event-outlook.md)|Occurs when the user selects the  **Send** action for an item (which is an instance of the parent object).|
|[Unload](postitem-unload-event-outlook.md)|Occurs before an Outlook item is unloaded from memory, either programmatically or by user action. |
|[Write](postitem-write-event-outlook.md)|Occurs when an instance of the parent object is saved, either explicitly (for example, using the  **[Save](postitem-save-method-outlook.md)** or **[SaveAs](postitem-saveas-method-outlook.md)** methods) or implicitly (for example, in response to a prompt when closing the item's inspector).|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[ClearConversationIndex](postitem-clearconversationindex-method-outlook.md)|Clears the index of the conversation thread for the post.|
|[ClearTaskFlag](postitem-cleartaskflag-method-outlook.md)|Clears the  **[PostItem](postitem-object-outlook.md)** object as a task.|
|[Close](postitem-close-method-outlook.md)|Closes and optionally saves changes to the Outlook item.|
|[Copy](postitem-copy-method-outlook.md)|Creates another instance of an object.|
|[Delete](postitem-delete-method-outlook.md)|Removes the item from the folder that contains the item.|
|[Display](postitem-display-method-outlook.md)|Displays a new  **[Inspector](inspector-object-outlook.md)** object for the item.|
|[Forward](postitem-forward-method-outlook.md)|Executes the  **Forward** action for an item and returns the resulting copy as a **[MailItem](mailitem-object-outlook.md)** object.|
|[GetConversation](postitem-getconversation-method-outlook.md)|Obtains a  **[Conversation](conversation-object-outlook.md)** object that represents the conversation to which this item belongs.|
|[MarkAsTask](postitem-markastask-method-outlook.md)|Marks a  **[PostItem](postitem-object-outlook.md)** object as a task and assigns a task interval for the object.|
|[Move](postitem-move-method-outlook.md)|Moves a Microsoft Outlook item to a new folder.|
|[Post](postitem-post-method-outlook.md)|Sends (posts) the  **[PostItem](postitem-object-outlook.md)** object.|
|[PrintOut](postitem-printout-method-outlook.md)|Prints the Outlook item using all default settings.The  **PrintOut** method is the only Outlook method that can be used for printing.|
|[Reply](postitem-reply-method-outlook.md)|Creates a reply, pre-addressed to the original sender, from the original message.|
|[Save](postitem-save-method-outlook.md)|Saves the Microsoft Outlook item to the current folder or, if this is a new item, to the Outlook default folder for the item type.|
|[SaveAs](postitem-saveas-method-outlook.md)|Saves the Microsoft Outlook item to the specified path and in the format of the specified file type. If the file type is not specified, the MSG format (.msg) is used.|
|[ShowCategoriesDialog](postitem-showcategoriesdialog-method-outlook.md)|Displays the  **Show Categories** dialog box, which allows you to select categories that correspond to the subject of the item.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Actions](postitem-actions-property-outlook.md)|Returns an  **[Actions](actions-object-outlook.md)** collection that represents all the available actions for the item. Read-only.|
|[Application](postitem-application-property-outlook.md)|Returns an  **[Application](application-object-outlook.md)** object that represents the parent Outlook application for the object. Read-only.|
|[Attachments](postitem-attachments-property-outlook.md)|Returns an  **[Attachments](attachments-object-outlook.md)** object that represents all the attachments for the specified item. Read-only.|
|[AutoResolvedWinner](postitem-autoresolvedwinner-property-outlook.md)|Returns a  **Boolean** that determines if the item is a winner of an automatic conflict resolution. Read-only.|
|[BillingInformation](postitem-billinginformation-property-outlook.md)|Returns or sets a  **String** representing the billing information associated with the Outlook item. Read/write.|
|[Body](postitem-body-property-outlook.md)|Returns or sets a  **String** representing the clear-text body of the Outlook item. Read/write.|
|[BodyFormat](postitem-bodyformat-property-outlook.md)|Returns or sets an  **[OlBodyFormat](olbodyformat-enumeration-outlook.md)** constant indicating the format of the body text. Read/write.|
|[Categories](postitem-categories-property-outlook.md)|Returns or sets a  **String** representing the categories assigned to the Outlook item. Read/write.|
|[Class](postitem-class-property-outlook.md)|Returns an  **[OlObjectClass](olobjectclass-enumeration-outlook.md)** constant indicating the object's class. Read-only.|
|[Companies](postitem-companies-property-outlook.md)|Returns or sets a  **String** representing the names of the companies associated with the Outlook item. Read/write.|
|[Conflicts](postitem-conflicts-property-outlook.md)|Return the  **[Conflicts](conflicts-object-outlook.md)** object that represents the items that are in conflict for any Outlook item object. Read-only.|
|[ConversationID](postitem-conversationid-property-outlook.md)|Returns a  **String** that uniquely identifies a **[Conversation](conversation-object-outlook.md)** object that the **[PostItem](postitem-object-outlook.md)** object belongs to. Read-only.|
|[ConversationIndex](postitem-conversationindex-property-outlook.md)|Returns a  **String** that indicates the relative position of the item within the conversation thread. Read-only.|
|[ConversationTopic](postitem-conversationtopic-property-outlook.md)|Returns a  **String** representing the topic of the conversation thread of the Outlook item. Read-only.|
|[CreationTime](postitem-creationtime-property-outlook.md)|Returns a  **Date** indicating the creation time for the Outlook item. Read-only.|
|[DownloadState](postitem-downloadstate-property-outlook.md)|Returns a constant that belongs to the  **[OlDownloadState](oldownloadstate-enumeration-outlook.md)** enumeration indicating the download state of the item. Read-only.|
|[EntryID](postitem-entryid-property-outlook.md)|Returns a  **String** representing the unique Entry ID of the object. Read-only.|
|[ExpiryTime](postitem-expirytime-property-outlook.md)|Returns or sets a  **Date** indicating the date and time at which the item becomes invalid and can be deleted. Read/write.|
|[FormDescription](postitem-formdescription-property-outlook.md)|Returns the  **[FormDescription](formdescription-object-outlook.md)** object that represents the form description for the specified Outlook item. Read-only.|
|[GetInspector](postitem-getinspector-property-outlook.md)|Returns an  **[Inspector](inspector-object-outlook.md)** object that represents an inspector initialized to contain the specified item. Read-only.|
|[HTMLBody](postitem-htmlbody-property-outlook.md)|Returns or sets a  **String** representing the HTML body of the specified item. Read/write.|
|[Importance](postitem-importance-property-outlook.md)|Returns or sets an  **[OlImportance](olimportance-enumeration-outlook.md)** constant indicating the relative importance level for the Outlook item. Read/write.|
|[InternetCodepage](postitem-internetcodepage-property-outlook.md)|Returns or sets a  **Long** that determines the Internet code page used by the item. Read/write.|
|[IsConflict](postitem-isconflict-property-outlook.md)|Returns a  **Boolean** that determines if the item is in conflict. Read-only.|
|[IsMarkedAsTask](postitem-ismarkedastask-property-outlook.md)|Returns a  **Boolean** value that indicates whether the **[PostItem](postitem-object-outlook.md)** is marked as a task. Read-only.|
|[ItemProperties](postitem-itemproperties-property-outlook.md)|Returns an  **[ItemProperties](itemproperties-object-outlook.md)** collection that represents all standard and user-defined properties associated with the Outlook item. Read-only.|
|[LastModificationTime](postitem-lastmodificationtime-property-outlook.md)|Returns a  **Date** specifying the date and time that the Outlook item was last modified. Read-only.|
|[MarkForDownload](postitem-markfordownload-property-outlook.md)|Returns or sets an  **[OlRemoteStatus](olremotestatus-enumeration-outlook.md)** constant that determines the status of an item once it is received by a remote user. Read/write.|
|[MessageClass](postitem-messageclass-property-outlook.md)|Returns or sets a  **String** representing the message class for the Outlook item. Read/write.|
|[Mileage](postitem-mileage-property-outlook.md)|Returns or sets a  **String** representing the mileage for an item. Read/write.|
|[NoAging](postitem-noaging-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** to not age the Outlook item. Read/write.|
|[OutlookInternalVersion](postitem-outlookinternalversion-property-outlook.md)|Returns a  **Long** representing the build number of the Outlook application for an Outlook item. Read-only.|
|[OutlookVersion](postitem-outlookversion-property-outlook.md)|Returns a  **String** indicating the major and minor version number of the Outlook application for an Outlook item. Read-only.|
|[Parent](postitem-parent-property-outlook.md)|Returns the parent  **Object** of the specified object. Read-only.|
|[PropertyAccessor](postitem-propertyaccessor-property-outlook.md)|Returns a  **[PropertyAccessor](propertyaccessor-object-outlook.md)** object that supports creating, getting, setting, and deleting properties of the parent **[PostItem](postitem-object-outlook.md)** object. Read-only.|
|[ReceivedTime](postitem-receivedtime-property-outlook.md)|Returns a  **Date** indicating the date and time at which the item was received. Read-only.|
|[ReminderOverrideDefault](postitem-reminderoverridedefault-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** if the reminder overrides the default reminder behavior for the item. Read/write.|
|[ReminderPlaySound](postitem-reminderplaysound-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** if the reminder should play a sound when it occurs for this item. Read/write.|
|[ReminderSet](postitem-reminderset-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** if a reminder has been set for this item. Read/write.|
|[ReminderSoundFile](postitem-remindersoundfile-property-outlook.md)|Returns or sets a  **String** indicating the path and file name of the sound file to play when the reminder occurs for the Outlook item. Read/write.|
|[ReminderTime](postitem-remindertime-property-outlook.md)|Returns or sets a  **Date** indicating the date and time at which the reminder should occur for the specified item. Read/write.|
|[RTFBody](postitem-rtfbody-property-outlook.md)|Returns or sets a  **Byte** array that represents the body of the Microsoft Outlook item in Rich Text Format. Read/write.|
|[Saved](postitem-saved-property-outlook.md)|Returns a  **Boolean** value that is **True** if the Outlook item has not been modified since the last save. Read-only.|
|[SenderEmailAddress](postitem-senderemailaddress-property-outlook.md)|Returns a  **String** that represents the e-mail address of the sender of the Outlook item. Read-only.|
|[SenderEmailType](postitem-senderemailtype-property-outlook.md)|Returns a  **String** that represents the type of entry for the e-mail address of the sender of the Outlook item, such as 'SMTP' for Internet address, 'EX' for a Microsoft Exchange server address, etc. Read-only.|
|[SenderName](postitem-sendername-property-outlook.md)|Returns a  **String** indicating the display name of the sender for the Outlook item. Read-only.|
|[Sensitivity](postitem-sensitivity-property-outlook.md)|Returns or sets a constant in the  **[OlSensitivity](olsensitivity-enumeration-outlook.md)** enumeration indicating the sensitivity for the Outlook item. Read/write.|
|[SentOn](postitem-senton-property-outlook.md)|Returns a  **Date** indicating the date and time on which the Outlook item was sent. Read-only.|
|[Session](postitem-session-property-outlook.md)|Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.|
|[Size](postitem-size-property-outlook.md)|Returns a  **Long** indicating the size (in bytes) of the Outlook item. Read-only.|
|[Subject](postitem-subject-property-outlook.md)|Returns or sets a  **String** indicating the subject for the Outlook item. Read/write.|
|[TaskCompletedDate](postitem-taskcompleteddate-property-outlook.md)|Returns or sets a  **Date** value that represents the completion date of the task for this **[PostItem](postitem-object-outlook.md)** . Read/write.|
|[TaskDueDate](postitem-taskduedate-property-outlook.md)|Returns or sets a  **Date** value that represents the due date of the task for this **[PostItem](postitem-object-outlook.md)** . Read/write.|
|[TaskStartDate](postitem-taskstartdate-property-outlook.md)|Returns or sets a  **Date** value that represents the start date of the task for this **[PostItem](postitem-object-outlook.md)** object. Read/write.|
|[TaskSubject](postitem-tasksubject-property-outlook.md)|Returns or sets a  **String** value that represents the subject of the task for the **[PostItem](postitem-object-outlook.md)** object. Read/write.|
|[ToDoTaskOrdinal](postitem-todotaskordinal-property-outlook.md)|Returns or sets a  **Date** value that represents the ordinal value of the task for the **[PostItem](postitem-object-outlook.md)** . Read/write.|
|[UnRead](postitem-unread-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** if the Outlook item has not been opened (read). Read/write.|
|[UserProperties](postitem-userproperties-property-outlook.md)|Returns the  **[UserProperties](userproperties-object-outlook.md)** collection that represents all the user properties for the Outlook item. Read-only.|

