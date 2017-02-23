---
title: TaskItem Members (Outlook)
ms.prod: OUTLOOK
ms.assetid: 97234a76-2fc5-bbe4-2e14-25ae18694fc9
---


# TaskItem Members (Outlook)
Represents a task (an assigned, delegated, or self-imposed task to be performed within a specified time frame) in a Tasks folder.

Represents a task (an assigned, delegated, or self-imposed task to be performed within a specified time frame) in a Tasks folder.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[AfterWrite](taskitem-afterwrite-event-outlook.md)|Occurs after Microsoft Outlook has saved the item.|
|[AttachmentAdd](taskitem-attachmentadd-event-outlook.md)|Occurs when an attachment has been added to an instance of the parent object.|
|[AttachmentRead](taskitem-attachmentread-event-outlook.md)|Occurs when an attachment in an instance of the parent object has been opened for reading.|
|[AttachmentRemove](taskitem-attachmentremove-event-outlook.md)|Occurs when an attachment has been removed from an instance of the parent object.|
|[BeforeAttachmentAdd](taskitem-beforeattachmentadd-event-outlook.md)|Occurs before an attachment is added to an instance of the parent object.|
|[BeforeAttachmentPreview](taskitem-beforeattachmentpreview-event-outlook.md)|Occurs before an attachment associated with an instance of the parent object is previewed.|
|[BeforeAttachmentRead](taskitem-beforeattachmentread-event-outlook.md)|Occurs before an attachment associated with an instance of the parent object is read from the file system, an attachment stream, or an  **[Attachment](attachment-object-outlook.md)** object.|
|[BeforeAttachmentSave](taskitem-beforeattachmentsave-event-outlook.md)|Occurs just before an attachment is saved.|
|[BeforeAttachmentWriteToTempFile](taskitem-beforeattachmentwritetotempfile-event-outlook.md)|Occurs before an attachment associated with an instance of the parent object is written to a temporary file.|
|[BeforeAutoSave](taskitem-beforeautosave-event-outlook.md)|Occurs before the item is automatically saved by Outlook.|
|[BeforeCheckNames](taskitem-beforechecknames-event-outlook.md)|Occurs just before Microsoft Outlook starts resolving names in the recipient collection for an item (which is an instance of the parent object).|
|[BeforeDelete](taskitem-beforedelete-event-outlook.md)|Occurs before an item (which is an instance of the parent object) is deleted.|
|[BeforeRead](taskitem-beforeread-event-outlook.md)|Occurs before Microsoft Outlook begins to read the properties for the item.|
|[Close](taskitem-close-event-outlook.md)|Occurs when the inspector associated with an item (which is an instance of the parent object) is being closed.|
|[CustomAction](taskitem-customaction-event-outlook.md)|Occurs when a custom action of an item (which is an instance of the parent object) executes.|
|[CustomPropertyChange](taskitem-custompropertychange-event-outlook.md)|Occurs when a custom property of an item (which is an instance of the parent object) is changed. |
|[Forward](taskitem-forward-event-outlook.md)|Occurs when the user selects the  **Forward** action for an item (which is an instance of the parent object).|
|[Open](taskitem-open-event-outlook.md)|Occurs when an instance of the parent object is being opened in an  **[Inspector](inspector-object-outlook.md)** .|
|[PropertyChange](taskitem-propertychange-event-outlook.md)|Occurs when an explicit built-in property (for example,  **[Subject](taskitem-subject-property-outlook.md)** ) of an instance of the parent object is changed.|
|[Read](taskitem-read-event-outlook.md)|Occurs when an instance of the parent object is opened for editing by the user. |
|[ReadComplete](taskitem-readcomplete-event-outlook.md)|Occurs when Outlook has completed reading the properties of the item.|
|[Reply](taskitem-reply-event-outlook.md)|Occurs when the user selects the  **Reply** action for an item (which is an instance of the parent object).|
|[ReplyAll](taskitem-replyall-event-outlook.md)|Occurs when the user selects the  **ReplyAll** action for an item (which is an instance of the parent object).|
|[Send](taskitem-send-event-outlook.md)|Occurs when the user selects the  **Send** action for an item, or when the **Send** method is called for the item, which is an instance of the parent object.|
|[Unload](taskitem-unload-event-outlook.md)|Occurs before an Outlook item is unloaded from memory, either programmatically or by user action. |
|[Write](taskitem-write-event-outlook.md)|Occurs when an instance of the parent object is saved, either explicitly (for example, using the  **[Save](taskitem-save-method-outlook.md)** or **[SaveAs](taskitem-saveas-method-outlook.md)** methods) or implicitly (for example, in response to a prompt when closing the item's inspector).|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Assign](taskitem-assign-method-outlook.md)|Assigns a task and returns a  **[TaskItem](taskitem-object-outlook.md)** object that represents it.|
|[CancelResponseState](taskitem-cancelresponsestate-method-outlook.md)|Resets an unsent response to a task request back to a simple task.|
|[ClearRecurrencePattern](taskitem-clearrecurrencepattern-method-outlook.md)|Removes the recurrence settings and restores the single-occurrence state for an appointment or task.|
|[Close](taskitem-close-method-outlook.md)|Closes and optionally saves changes to the Outlook item.|
|[Copy](taskitem-copy-method-outlook.md)|Creates another instance of an object.|
|[Delete](taskitem-delete-method-outlook.md)|Removes the item from the folder that contains the item.|
|[Display](taskitem-display-method-outlook.md)|Displays a new  **[Inspector](inspector-object-outlook.md)** object for the item.|
|[GetConversation](taskitem-getconversation-method-outlook.md)|Obtains a  **[Conversation](conversation-object-outlook.md)** object that represents the conversation to which this item belongs.|
|[GetRecurrencePattern](taskitem-getrecurrencepattern-method-outlook.md)|Returns a  **[RecurrencePattern](recurrencepattern-object-outlook.md)** object that represents the recurrence attributes of a task. If there is no existing recurrence pattern, a new empty **RecurrencePattern** object is returned.|
|[MarkComplete](taskitem-markcomplete-method-outlook.md)|Marks the task as completed.|
|[Move](taskitem-move-method-outlook.md)|Moves a Microsoft Outlook item to a new folder.|
|[PrintOut](taskitem-printout-method-outlook.md)|Prints the Outlook item using all default settings.The  **PrintOut** method is the only Outlook method that can be used for printing.|
|[Respond](taskitem-respond-method-outlook.md)|Responds to a task request.|
|[Save](taskitem-save-method-outlook.md)|Saves the Microsoft Outlook item to the current folder or, if this is a new item, to the Outlook default folder for the item type.|
|[SaveAs](taskitem-saveas-method-outlook.md)|Saves the Microsoft Outlook item to the specified path and in the format of the specified file type. If the file type is not specified, the MSG format (.msg) is used.|
|[Send](taskitem-send-method-outlook.md)|Sends the task.|
|[ShowCategoriesDialog](taskitem-showcategoriesdialog-method-outlook.md)|Displays the  **Show Categories** dialog box, which allows you to select categories that correspond to the subject of the item.|
|[SkipRecurrence](taskitem-skiprecurrence-method-outlook.md)|Clears the current instance of a recurring task and sets the recurrence to the next instance of that task.|
|[StatusReport](taskitem-statusreport-method-outlook.md)|Sends a status report to all Cc recipients (recipients returned by the  **[StatusUpdateRecipients](taskitem-statusupdaterecipients-property-outlook.md)** property) with the current status for the task and returns an **Object** representing the status report.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Actions](taskitem-actions-property-outlook.md)|Returns an  **[Actions](actions-object-outlook.md)** collection that represents all the available actions for the item. Read-only.|
|[ActualWork](taskitem-actualwork-property-outlook.md)|Returns or sets a  **Long** indicating the actual effort spent on the task. Read/write.|
|[Application](taskitem-application-property-outlook.md)|Returns an  **[Application](application-object-outlook.md)** object that represents the parent Outlook application for the object. Read-only.|
|[Attachments](taskitem-attachments-property-outlook.md)|Returns an  **[Attachments](attachments-object-outlook.md)** object that represents all the attachments for the specified item. Read-only.|
|[AutoResolvedWinner](taskitem-autoresolvedwinner-property-outlook.md)|Returns a  **Boolean** that determines if the item is a winner of an automatic conflict resolution. Read-only.|
|[BillingInformation](taskitem-billinginformation-property-outlook.md)|Returns or sets a  **String** representing the billing information associated with the Outlook item. Read/write.|
|[Body](taskitem-body-property-outlook.md)|Returns or sets a  **String** representing the clear-text body of the Outlook item. Read/write.|
|[CardData](taskitem-carddata-property-outlook.md)|Returns or sets a  **String** representing the text of the card data for the task. Read/write.|
|[Categories](taskitem-categories-property-outlook.md)|Returns or sets a  **String** representing the categories assigned to the Outlook item. Read/write.|
|[Class](taskitem-class-property-outlook.md)|Returns an  **[OlObjectClass](olobjectclass-enumeration-outlook.md)** constant indicating the object's class. Read-only.|
|[Companies](taskitem-companies-property-outlook.md)|Returns or sets a  **String** representing the names of the companies associated with the Outlook item. Read/write.|
|[Complete](taskitem-complete-property-outlook.md)|Returns a  **Boolean** value that indicates whether the task is completed. Read/write **Boolean** .|
|[Conflicts](taskitem-conflicts-property-outlook.md)|Return the  **[Conflicts](conflicts-object-outlook.md)** object that represents the items that are in conflict for any Outlook item object. Read-only.|
|[ContactNames](taskitem-contactnames-property-outlook.md)|Returns or sets a  **String** representing the contact names associated with the Outlook item. Read/write.|
|[ConversationID](taskitem-conversationid-property-outlook.md)|Returns a  **String** that uniquely identifies a **[Conversation](conversation-object-outlook.md)** object that the **[TaskItem](taskitem-object-outlook.md)** object belongs to. Read-only.|
|[ConversationIndex](taskitem-conversationindex-property-outlook.md)|Returns a  **String** that indicates the relative position of the item within the conversation thread. Read-only.|
|[ConversationTopic](taskitem-conversationtopic-property-outlook.md)|Returns a  **String** representing the topic of the conversation thread of the Outlook item. Read-only.|
|[CreationTime](taskitem-creationtime-property-outlook.md)|Returns a  **Date** indicating the creation time for the Outlook item. Read-only.|
|[DateCompleted](taskitem-datecompleted-property-outlook.md)|Returns or sets a  **Date** indicating the completion date of the task. Read/write.|
|[DelegationState](taskitem-delegationstate-property-outlook.md)|Returns an  **[OlTaskDelegationState](oltaskdelegationstate-enumeration-outlook.md)** constant indicating the delegation state of the task. Read-only.|
|[Delegator](taskitem-delegator-property-outlook.md)|Returns a  **String** representing the display name of the delegator for the task. Read-only.|
|[DownloadState](taskitem-downloadstate-property-outlook.md)|Returns a constant that belongs to the  **[OlDownloadState](oldownloadstate-enumeration-outlook.md)** enumeration indicating the download state of the item. Read-only.|
|[DueDate](taskitem-duedate-property-outlook.md)|Returns or sets a  **Date** indicating the due date for the task. Read/write.|
|[EntryID](taskitem-entryid-property-outlook.md)|Returns a  **String** representing the unique Entry ID of the object. Read-only.|
|[FormDescription](taskitem-formdescription-property-outlook.md)|Returns the  **[FormDescription](formdescription-object-outlook.md)** object that represents the form description for the specified Outlook item. Read-only.|
|[GetInspector](taskitem-getinspector-property-outlook.md)|Returns an  **[Inspector](inspector-object-outlook.md)** object that represents an inspector initialized to contain the specified item. Read-only.|
|[Importance](taskitem-importance-property-outlook.md)|Returns or sets an  **[OlImportance](olimportance-enumeration-outlook.md)** constant indicating the relative importance level for the Outlook item. Read/write.|
|[InternetCodepage](taskitem-internetcodepage-property-outlook.md)|Returns or sets a  **Long** that determines the Internet code page used by the item. Read/write.|
|[IsConflict](taskitem-isconflict-property-outlook.md)|Returns a  **Boolean** that determines if the item is in conflict. Read-only.|
|[IsRecurring](taskitem-isrecurring-property-outlook.md)|Returns a  **Boolean** value that is **True** if the task is a recurring task. Read-only.|
|[ItemProperties](taskitem-itemproperties-property-outlook.md)|Returns an  **[ItemProperties](itemproperties-object-outlook.md)** collection that represents all standard and user-defined properties associated with the Outlook item. Read-only.|
|[LastModificationTime](taskitem-lastmodificationtime-property-outlook.md)|Returns a  **Date** specifying the date and time that the Outlook item was last modified. Read-only.|
|[MarkForDownload](taskitem-markfordownload-property-outlook.md)|Returns or sets an  **[OlRemoteStatus](olremotestatus-enumeration-outlook.md)** constant that determines the status of an item once it is received by a remote user. Read/write.|
|[MessageClass](taskitem-messageclass-property-outlook.md)|Returns or sets a  **String** representing the message class for the Outlook item. Read/write.|
|[Mileage](taskitem-mileage-property-outlook.md)|Returns or sets a  **String** representing the mileage for an item. Read/write.|
|[NoAging](taskitem-noaging-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** to not age the Outlook item. Read/write.|
|[Ordinal](taskitem-ordinal-property-outlook.md)|Returns or sets a  **Long** specifying the position in the view (ordinal) for the task. Read/write.|
|[OutlookInternalVersion](taskitem-outlookinternalversion-property-outlook.md)|Returns a  **Long** representing the build number of the Outlook application for an Outlook item. Read-only.|
|[OutlookVersion](taskitem-outlookversion-property-outlook.md)|Returns a  **String** indicating the major and minor version number of the Outlook application for an Outlook item. Read-only.|
|[Owner](taskitem-owner-property-outlook.md)|Returns or sets a  **String** indicating the owner for the task.|
|[Ownership](taskitem-ownership-property-outlook.md)|Returns an  **[OlTaskOwnership](oltaskownership-enumeration-outlook.md)** specifying the ownership state of the task. Read-only.|
|[Parent](taskitem-parent-property-outlook.md)|Returns the parent  **Object** of the specified object. Read-only.|
|[PercentComplete](taskitem-percentcomplete-property-outlook.md)|Returns or sets a  **Long** indicating the percentage of the task completed at the current date and time. Read/write.|
|[PropertyAccessor](taskitem-propertyaccessor-property-outlook.md)|Returns a  **[PropertyAccessor](propertyaccessor-object-outlook.md)** object that supports creating, getting, setting, and deleting properties of the parent **[TaskItem](taskitem-object-outlook.md)** object. Read-only.|
|[Recipients](taskitem-recipients-property-outlook.md)|Returns a  **[Recipients](recipients-object-outlook.md)** collection that represents all the recipients for the Outlook item. Read-only.|
|[ReminderOverrideDefault](taskitem-reminderoverridedefault-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** if the reminder overrides the default reminder behavior for the item. Read/write.|
|[ReminderPlaySound](taskitem-reminderplaysound-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** if the reminder should play a sound when it occurs for this item. Read/write.|
|[ReminderSet](taskitem-reminderset-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** if a reminder has been set for this item. Read/write.|
|[ReminderSoundFile](taskitem-remindersoundfile-property-outlook.md)|Returns or sets a  **String** indicating the path and file name of the sound file to play when the reminder occurs for the Outlook item. Read/write.|
|[ReminderTime](taskitem-remindertime-property-outlook.md)|Returns or sets a  **Date** indicating the date and time at which the reminder should occur for the specified item. Read/write.|
|[ResponseState](taskitem-responsestate-property-outlook.md)|Returns an  **[OlTaskResponse](oltaskresponse-enumeration-outlook.md)** constant indicating the overall status of the response to the specified task request. Read-only.|
|[Role](taskitem-role-property-outlook.md)|Returns or sets a  **String** containing the free-form text string associating the owner of a task with a role for the task. Read/write.|
|[RTFBody](taskitem-rtfbody-property-outlook.md)|Returns or sets a  **Byte** array that represents the body of the Microsoft Outlook item in Rich Text Format. Read/write.|
|[Saved](taskitem-saved-property-outlook.md)|Returns a  **Boolean** value that is **True** if the Outlook item has not been modified since the last save. Read-only.|
|[SchedulePlusPriority](taskitem-schedulepluspriority-property-outlook.md)|Returns or sets a  **String** representing the Microsoft Schedule+ priority for the task. Read/write.|
|[SendUsingAccount](taskitem-sendusingaccount-property-outlook.md)|Returns or sets an  **[Account](account-object-outlook.md)** object that represents the account under which the **[TaskItem](taskitem-object-outlook.md)** object is to be sent. Read/write.|
|[Sensitivity](taskitem-sensitivity-property-outlook.md)|Returns or sets a constant in the  **[OlSensitivity](olsensitivity-enumeration-outlook.md)** enumeration indicating the sensitivity for the Outlook item. Read/write.|
|[Session](taskitem-session-property-outlook.md)|Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.|
|[Size](taskitem-size-property-outlook.md)|Returns a  **Long** indicating the size (in bytes) of the Outlook item. Read-only.|
|[StartDate](taskitem-startdate-property-outlook.md)|Returns or sets a  **Date** indicating the start date for the task. Read/write.|
|[Status](taskitem-status-property-outlook.md)|Returns or sets an  **[OlTaskStatus](oltaskstatus-enumeration-outlook.md)** constant specifying the status for the task. Corresponds to the **Status** field of a **[TaskItem](taskitem-object-outlook.md)** . Read/write.|
|[StatusOnCompletionRecipients](taskitem-statusoncompletionrecipients-property-outlook.md)|Returns or sets a semicolon-delimited  **String** of display names for recipients who will receive status upon completion of the task. Read/write.|
|[StatusUpdateRecipients](taskitem-statusupdaterecipients-property-outlook.md)|Returns a semicolon-delimited  **String** of display names for recipients who receive status updates for the task. Read/write.|
|[Subject](taskitem-subject-property-outlook.md)|Returns or sets a  **String** indicating the subject for the Outlook item. Read/write.|
|[TeamTask](taskitem-teamtask-property-outlook.md)|Returns a  **Boolean** that indicates **True** if the task is a team task. Read/write.|
|[ToDoTaskOrdinal](taskitem-todotaskordinal-property-outlook.md)|Returns or sets a  **Date** value that represents the ordinal value of the task for the **[TaskItem](taskitem-object-outlook.md)** . Read/write.|
|[TotalWork](taskitem-totalwork-property-outlook.md)|Returns or sets a  **Long** indicating the total work for the task. Read/write.|
|[UnRead](taskitem-unread-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** if the Outlook item has not been opened (read). Read/write.|
|[UserProperties](taskitem-userproperties-property-outlook.md)|Returns the  **[UserProperties](userproperties-object-outlook.md)** collection that represents all the user properties for the Outlook item. Read-only.|

