---
title: MailItem Members (Outlook)
ms.prod: OUTLOOK
ms.assetid: 1094d7df-ee80-a4b0-5a21-db2979506e6b
---


# MailItem Members (Outlook)
Represents a mail message.

Represents a mail message.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[AfterWrite](mailitem-afterwrite-event-outlook.md)|Occurs after Microsoft Outlook has saved the item.|
|[AttachmentAdd](mailitem-attachmentadd-event-outlook.md)|Occurs when an attachment has been added to an instance of the parent object.|
|[AttachmentRead](mailitem-attachmentread-event-outlook.md)|Occurs when an attachment in an instance of the parent object has been opened for reading.|
|[AttachmentRemove](mailitem-attachmentremove-event-outlook.md)|Occurs when an attachment has been removed from an instance of the parent object.|
|[BeforeAttachmentAdd](mailitem-beforeattachmentadd-event-outlook.md)|Occurs before an attachment is added to an instance of the parent object.|
|[BeforeAttachmentPreview](mailitem-beforeattachmentpreview-event-outlook.md)|Occurs before an attachment associated with an instance of the parent object is previewed.|
|[BeforeAttachmentRead](mailitem-beforeattachmentread-event-outlook.md)|Occurs before an attachment associated with an instance of the parent object is read from the file system, an attachment stream, or an  **[Attachment](attachment-object-outlook.md)** object.|
|[BeforeAttachmentSave](mailitem-beforeattachmentsave-event-outlook.md)|Occurs just before an attachment is saved.|
|[BeforeAttachmentWriteToTempFile](mailitem-beforeattachmentwritetotempfile-event-outlook.md)|Occurs before an attachment associated with an instance of the parent object is written to a temporary file.|
|[BeforeAutoSave](mailitem-beforeautosave-event-outlook.md)|Occurs before the item is automatically saved by Outlook.|
|[BeforeCheckNames](mailitem-beforechecknames-event-outlook.md)|Occurs just before Microsoft Outlook starts resolving names in the recipient collection for an item (which is an instance of the parent object).|
|[BeforeDelete](mailitem-beforedelete-event-outlook.md)|Occurs before an item (which is an instance of the parent object) is deleted.|
|[BeforeRead](mailitem-beforeread-event-outlook.md)|Occurs before Microsoft Outlook begins to read the properties for the item.|
|[Close](mailitem-close-event-outlook.md)|Occurs when the inspector associated with an item (which is an instance of the parent object) is being closed.|
|[CustomAction](mailitem-customaction-event-outlook.md)|Occurs when a custom action of an item (which is an instance of the parent object) executes.|
|[CustomPropertyChange](mailitem-custompropertychange-event-outlook.md)|Occurs when a custom property of an item (which is an instance of the parent object) is changed. |
|[Forward](mailitem-forward-event-outlook.md)|Occurs when the user selects the  **Forward** action for an item, or when the **Forward** method is called for the item, which is an instance of the parent object.|
|[Open](mailitem-open-event-outlook.md)|Occurs when an instance of the parent object is being opened in an  **[Inspector](inspector-object-outlook.md)** .|
|[PropertyChange](mailitem-propertychange-event-outlook.md)|Occurs when an explicit built-in property (for example,  **[Subject](appointmentitem-subject-property-outlook.md)** ) of an instance of the parent object is changed.|
|[Read](mailitem-read-event-outlook.md)|Occurs when an instance of the parent object is opened for editing by the user. |
|[ReadComplete](mailitem-readcomplete-event-outlook.md)|Occurs when Outlook has completed reading the properties of the item.|
|[Reply](mailitem-reply-event-outlook.md)|Occurs when the user selects the  **Reply** action for an item, or when the **Reply** method is called for the item, which is an instance of the parent object.|
|[ReplyAll](mailitem-replyall-event-outlook.md)|Occurs when the user selects the  **ReplyAll** action for an item, or when the **ReplyAll** method is called for the item, which is an instance of the parent object.|
|[Send](mailitem-send-event-outlook.md)|Occurs when the user selects the  **Send** action for an item, or when the **Send** method is called for the item, which is an instance of the parent object.|
|[Unload](mailitem-unload-event-outlook.md)|Occurs before an Outlook item is unloaded from memory, either programmatically or by user action. |
|[Write](mailitem-write-event-outlook.md)|Occurs when an instance of the parent object is saved, either explicitly (for example, using the  **[Save](mailitem-save-method-outlook.md)** or **[SaveAs](mailitem-saveas-method-outlook.md)** methods) or implicitly (for example, in response to a prompt when closing the item's inspector).|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[AddBusinessCard](mailitem-addbusinesscard-method-outlook.md)|Appends contact information based on the Electronic Business Card (EBC) associated with the specified  **[ContactItem](contactitem-object-outlook.md)** object to the **[MailItem](mailitem-object-outlook.md)** object.|
|[ClearConversationIndex](mailitem-clearconversationindex-method-outlook.md)|Clears the index of the conversation thread for the mail message.|
|[ClearTaskFlag](mailitem-cleartaskflag-method-outlook.md)|Clears the  **[MailItem](mailitem-object-outlook.md)** object as a task.|
|[Close](mailitem-close-method-outlook.md)|Closes and optionally saves changes to the Outlook item.|
|[Copy](mailitem-copy-method-outlook.md)|Creates another instance of an object.|
|[Delete](mailitem-delete-method-outlook.md)|Removes the item from the folder that contains the item.|
|[Display](mailitem-display-method-outlook.md)|Displays a new  **[Inspector](inspector-object-outlook.md)** object for the item.|
|[Forward](mailitem-forward-method-outlook.md)|Executes the  **Forward** action for an item and returns the resulting copy as a **[MailItem](mailitem-object-outlook.md)** object.|
|[GetConversation](mailitem-getconversation-method-outlook.md)|Obtains a  **[Conversation](conversation-object-outlook.md)** object that represents the conversation to which this item belongs.|
|[MarkAsTask](mailitem-markastask-method-outlook.md)|Marks a  **[MailItem](mailitem-object-outlook.md)** object as a task and assigns a task interval for the object.|
|[Move](mailitem-move-method-outlook.md)|Moves a Microsoft Outlook item to a new folder.|
|[PrintOut](mailitem-printout-method-outlook.md)|Prints the Outlook item using all default settings.The  **PrintOut** method is the only Outlook method that can be used for printing.|
|[Reply](mailitem-reply-method-outlook.md)|Creates a reply, pre-addressed to the original sender, from the original message.|
|[ReplyAll](mailitem-replyall-method-outlook.md)|Creates a reply to all original recipients from the original message.|
|[Save](mailitem-save-method-outlook.md)|Saves the Microsoft Outlook item to the current folder or, if this is a new item, to the Outlook default folder for the item type.|
|[SaveAs](mailitem-saveas-method-outlook.md)|Saves the Microsoft Outlook item to the specified path and in the format of the specified file type. If the file type is not specified, the MSG format (.msg) is used.|
|[Send](mailitem-send-method-outlook.md)|Sends the e-mail message.|
|[ShowCategoriesDialog](mailitem-showcategoriesdialog-method-outlook.md)|Displays the  **Show Categories** dialog box, which allows you to select categories that correspond to the subject of the item.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Actions](mailitem-actions-property-outlook.md)|Returns an  **[Actions](actions-object-outlook.md)** collection that represents all the available actions for the item. Read-only.|
|[AlternateRecipientAllowed](mailitem-alternaterecipientallowed-property-outlook.md)|Returns a  **Boolean** value that indicates whether the mail message can be forwarded. Read/write.|
|[Application](mailitem-application-property-outlook.md)|Returns an  **[Application](application-object-outlook.md)** object that represents the parent Outlook application for the object. Read-only.|
|[Attachments](mailitem-attachments-property-outlook.md)|Returns an  **[Attachments](attachments-object-outlook.md)** object that represents all the attachments for the specified item. Read-only.|
|[AutoForwarded](mailitem-autoforwarded-property-outlook.md)|A  **Boolean** value that returns **True** if the item was automatically forwarded. Read/write.|
|[AutoResolvedWinner](mailitem-autoresolvedwinner-property-outlook.md)|Returns a  **Boolean** that determines if the item is a winner of an automatic conflict resolution. Read-only.|
|[BCC](mailitem-bcc-property-outlook.md)|Returns a  **String** representing the display list of blind carbon copy (BCC) names for a **[MailItem](mailitem-object-outlook.md)** . Read/write.|
|[BillingInformation](mailitem-billinginformation-property-outlook.md)|Returns or sets a  **String** representing the billing information associated with the Outlook item. Read/write.|
|[Body](mailitem-body-property-outlook.md)|Returns or sets a  **String** representing the clear-text body of the Outlook item. Read/write.|
|[BodyFormat](mailitem-bodyformat-property-outlook.md)|Returns or sets an  **[OlBodyFormat](olbodyformat-enumeration-outlook.md)** constant indicating the format of the body text. Read/write.|
|[Categories](mailitem-categories-property-outlook.md)|Returns or sets a  **String** representing the categories assigned to the Outlook item. Read/write.|
|[CC](mailitem-cc-property-outlook.md)|Returns a  **String** representing the display list of carbon copy (CC) names for a **[MailItem](mailitem-object-outlook.md)** . Read/write.|
|[Class](mailitem-class-property-outlook.md)|Returns an  **[OlObjectClass](olobjectclass-enumeration-outlook.md)** constant indicating the object's class. Read-only.|
|[Companies](mailitem-companies-property-outlook.md)|Returns or sets a  **String** representing the names of the companies associated with the Outlook item. Read/write.|
|[Conflicts](mailitem-conflicts-property-outlook.md)|Return the  **[Conflicts](conflicts-object-outlook.md)** object that represents the items that are in conflict for any Outlook item object. Read-only.|
|[ConversationID](mailitem-conversationid-property-outlook.md)|Returns a  **String** that uniquely identifies a **[Conversation](conversation-object-outlook.md)** object that the **[MailItem](mailitem-object-outlook.md)** object belongs to. Read-only.|
|[ConversationIndex](mailitem-conversationindex-property-outlook.md)|Returns a  **String** that indicates the relative position of the item within the conversation thread. Read-only.|
|[ConversationTopic](mailitem-conversationtopic-property-outlook.md)|Returns a  **String** representing the topic of the conversation thread of the Outlook item. Read-only.|
|[CreationTime](mailitem-creationtime-property-outlook.md)|Returns a  **Date** indicating the creation time for the Outlook item. Read-only.|
|[DeferredDeliveryTime](mailitem-deferreddeliverytime-property-outlook.md)|Returns or sets a  **Date** indicating the date and time the mail message is to be delivered. Read/write.|
|[DeleteAfterSubmit](mailitem-deleteaftersubmit-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** if a copy of the mail message is not saved upon being sent, and **False** if a copy is saved. Read/write.|
|[DownloadState](mailitem-downloadstate-property-outlook.md)|Returns a constant that belongs to the  **[OlDownloadState](oldownloadstate-enumeration-outlook.md)** enumeration indicating the download state of the item. Read-only.|
|[EntryID](mailitem-entryid-property-outlook.md)|Returns a  **String** representing the unique Entry ID of the object. Read-only.|
|[ExpiryTime](mailitem-expirytime-property-outlook.md)|Returns or sets a  **Date** indicating the date and time at which the item becomes invalid and can be deleted. Read/write.|
|[FlagRequest](mailitem-flagrequest-property-outlook.md)|Returns or sets a  **String** that indicates the requested action for a mail item. Read/write.|
|[FormDescription](mailitem-formdescription-property-outlook.md)|Returns the  **[FormDescription](formdescription-object-outlook.md)** object that represents the form description for the specified Outlook item. Read-only.|
|[GetInspector](mailitem-getinspector-property-outlook.md)|Returns an  **[Inspector](inspector-object-outlook.md)** object that represents an inspector initialized to contain the specified item. Read-only.|
|[HTMLBody](mailitem-htmlbody-property-outlook.md)|Returns or sets a  **String** representing the HTML body of the specified item. Read/write.|
|[Importance](mailitem-importance-property-outlook.md)|Returns or sets an  **[OlImportance](olimportance-enumeration-outlook.md)** constant indicating the relative importance level for the Outlook item. Read/write.|
|[InternetCodepage](mailitem-internetcodepage-property-outlook.md)|Returns or sets a  **Long** that determines the Internet code page used by the item. Read/write.|
|[IsConflict](mailitem-isconflict-property-outlook.md)|Returns a  **Boolean** that determines if the item is in conflict. Read-only.|
|[IsMarkedAsTask](mailitem-ismarkedastask-property-outlook.md)|Returns a  **Boolean** value that indicates whether the **[MailItem](mailitem-object-outlook.md)** is marked as a task. Read-only.|
|[ItemProperties](mailitem-itemproperties-property-outlook.md)|Returns an  **[ItemProperties](itemproperties-object-outlook.md)** collection that represents all standard and user-defined properties associated with the Outlook item. Read-only.|
|[LastModificationTime](mailitem-lastmodificationtime-property-outlook.md)|Returns a  **Date** specifying the date and time that the Outlook item was last modified. Read-only.|
|[MarkForDownload](mailitem-markfordownload-property-outlook.md)|Returns or sets an  **[OlRemoteStatus](olremotestatus-enumeration-outlook.md)** constant that determines the status of an item once it is received by a remote user. Read/write.|
|[MessageClass](mailitem-messageclass-property-outlook.md)|Returns or sets a  **String** representing the message class for the Outlook item. Read/write.|
|[Mileage](mailitem-mileage-property-outlook.md)|Returns or sets a  **String** representing the mileage for an item. Read/write.|
|[NoAging](mailitem-noaging-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** to not age the Outlook item. Read/write.|
|[OriginatorDeliveryReportRequested](mailitem-originatordeliveryreportrequested-property-outlook.md)|Returns or sets a  **Boolean** value that determines whether the originator of the meeting item or mail message will receive a delivery report. Read/write.|
|[OutlookInternalVersion](mailitem-outlookinternalversion-property-outlook.md)|Returns a  **Long** representing the build number of the Outlook application for an Outlook item. Read-only.|
|[OutlookVersion](mailitem-outlookversion-property-outlook.md)|Returns a  **String** indicating the major and minor version number of the Outlook application for an Outlook item. Read-only.|
|[Parent](mailitem-parent-property-outlook.md)|Returns the parent  **Object** of the specified object. Read-only.|
|[Permission](mailitem-permission-property-outlook.md)|Sets or returns an  **[OlPermission](olpermission-enumeration-outlook.md)** constant that determines what permissions to grant to the recipients of the e-mail item. Read/write.|
|[PermissionService](mailitem-permissionservice-property-outlook.md)|Sets or returns an  **[OlPermissionService](olpermissionservice-enumeration-outlook.md)** constant that determines the permission service that will be used when sending a message protected by Information Rights Management (IRM). Read/write.|
|[PermissionTemplateGuid](mailitem-permissiontemplateguid-property-outlook.md)|Returns or sets a  **String** value that represents the GUID of the template file to apply to the **[MailItem](mailitem-object-outlook.md)** in order to specify Information Rights Management (IRM) permissions. Read/write.|
|[PropertyAccessor](mailitem-propertyaccessor-property-outlook.md)|Returns a  **[PropertyAccessor](propertyaccessor-object-outlook.md)** object that supports creating, getting, setting, and deleting properties of the parent **[MailItem](mailitem-object-outlook.md)** object. Read-only.|
|[ReadReceiptRequested](mailitem-readreceiptrequested-property-outlook.md)|Returns a  **Boolean** value that indicates **True** if a read receipt has been requested by the sender.|
|[ReceivedByEntryID](mailitem-receivedbyentryid-property-outlook.md)|Returns a  **String** representing the **[EntryID](recipient-entryid-property-outlook.md)** for the true recipient as set by the transport provider delivering the mail message. Read-only.|
|[ReceivedByName](mailitem-receivedbyname-property-outlook.md)|Returns a  **String** representing the display name of the true recipient for the mail message. Read-only.|
|[ReceivedOnBehalfOfEntryID](mailitem-receivedonbehalfofentryid-property-outlook.md)|Returns a  **String** representing the **[EntryID](recipient-entryid-property-outlook.md)** of the user delegated to represent the recipient for the mail message. Read-only.|
|[ReceivedOnBehalfOfName](mailitem-receivedonbehalfofname-property-outlook.md)|Returns a  **String** representing the display name of the user delegated to represent the recipient for the mail message. Read-only.|
|[ReceivedTime](mailitem-receivedtime-property-outlook.md)|Returns a  **Date** indicating the date and time at which the item was received. Read-only.|
|[RecipientReassignmentProhibited](mailitem-recipientreassignmentprohibited-property-outlook.md)|Returns a  **Boolean** that indicates **True** if the recipient cannot forward the mail message. Read/write.|
|[Recipients](mailitem-recipients-property-outlook.md)|Returns a  **[Recipients](recipients-object-outlook.md)** collection that represents all the recipients for the Outlook item. Read-only.|
|[ReminderOverrideDefault](mailitem-reminderoverridedefault-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** if the reminder overrides the default reminder behavior for the item. Read/write.|
|[ReminderPlaySound](mailitem-reminderplaysound-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** if the reminder should play a sound when it occurs for this item. Read/write.|
|[ReminderSet](mailitem-reminderset-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** if a reminder has been set for this item. Read/write.|
|[ReminderSoundFile](mailitem-remindersoundfile-property-outlook.md)|Returns or sets a  **String** indicating the path and file name of the sound file to play when the reminder occurs for the Outlook item. Read/write.|
|[ReminderTime](mailitem-remindertime-property-outlook.md)|Returns or sets a  **Date** indicating the date and time at which the reminder should occur for the specified item. Read/write.|
|[RemoteStatus](mailitem-remotestatus-property-outlook.md)|Returns or sets an  **[OlRemoteStatus](olremotestatus-enumeration-outlook.md)** constant specifying the remote status of the mail message. Read/write.|
|[ReplyRecipientNames](mailitem-replyrecipientnames-property-outlook.md)|Returns a semicolon-delimited  **String** list of reply recipients for the mail message. Read-only.|
|[ReplyRecipients](mailitem-replyrecipients-property-outlook.md)|Returns a  **[Recipients](recipients-object-outlook.md)** collection that represents all the reply recipient objects for the Outlook item. Read-only.|
|[RetentionExpirationDate](mailitem-retentionexpirationdate-property-outlook.md)|Returns a  **Date** that specifies the date when the **[MailItem](mailitem-object-outlook.md)** object expires, after which the Messaging Records Management (MRM) Assistant will delete the item. Read-only.|
|[RetentionPolicyName](mailitem-retentionpolicyname-property-outlook.md)|Returns a  **String** that specifies the name of the retention policy. Read-only.|
|[RTFBody](mailitem-rtfbody-property-outlook.md)|Returns or sets a  **Byte** array that represents the body of the Microsoft Outlook item in Rich Text Format. Read/write.|
|[Saved](mailitem-saved-property-outlook.md)|Returns a  **Boolean** value that is **True** if the Outlook item has not been modified since the last save. Read-only.|
|[SaveSentMessageFolder](mailitem-savesentmessagefolder-property-outlook.md)|Returns or sets a  **[Folder](folder-object-outlook.md)** object that represents the folder in which a copy of the e-mail message will be saved after being sent. Read/write.|
|[Sender](mailitem-sender-property-outlook.md)|Returns or sets an [AddressEntry](addressentry-object-outlook.md) object that corresponds to the user of the account from which the[MailItem](mailitem-object-outlook.md) is sent. Read/write.|
|[SenderEmailAddress](mailitem-senderemailaddress-property-outlook.md)|Returns a  **String** that represents the e-mail address of the sender of the Outlook item. Read-only.|
|[SenderEmailType](mailitem-senderemailtype-property-outlook.md)|Returns a  **String** that represents the type of entry for the e-mail address of the sender of the Outlook item, such as 'SMTP' for Internet address, 'EX' for a Microsoft Exchange server address, etc. Read-only.|
|[SenderName](mailitem-sendername-property-outlook.md)|Returns a  **String** indicating the display name of the sender for the Outlook item. Read-only.|
|[SendUsingAccount](mailitem-sendusingaccount-property-outlook.md)|Returns or sets an  **[Account](account-object-outlook.md)** object that represents the account under which the **[MailItem](mailitem-object-outlook.md)** is to be sent. Read/write.|
|[Sensitivity](mailitem-sensitivity-property-outlook.md)|Returns or sets a constant in the  **[OlSensitivity](olsensitivity-enumeration-outlook.md)** enumeration indicating the sensitivity for the Outlook item. Read/write.|
|[Sent](mailitem-sent-property-outlook.md)|Returns a  **Boolean** value that indicates if a message has been sent. Read-only.|
|[SentOn](mailitem-senton-property-outlook.md)|Returns a  **Date** indicating the date and time on which the Outlook item was sent. Read-only.|
|[SentOnBehalfOfName](mailitem-sentonbehalfofname-property-outlook.md)|Returns a  **String** indicating the display name for the intended sender of the mail message. Read/write.|
|[Session](mailitem-session-property-outlook.md)|Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.|
|[Size](mailitem-size-property-outlook.md)|Returns a  **Long** indicating the size (in bytes) of the Outlook item. Read-only.|
|[Subject](mailitem-subject-property-outlook.md)|Returns or sets a  **String** indicating the subject for the Outlook item. Read/write.|
|[Submitted](mailitem-submitted-property-outlook.md)|Returns a  **Boolean** value that is **True** if the item has been submitted. Read-only.|
|[TaskCompletedDate](mailitem-taskcompleteddate-property-outlook.md)|Returns or sets a  **Date** value that represents the completion date of the task for this **[MailItem](mailitem-object-outlook.md)** . Read/write.|
|[TaskDueDate](mailitem-taskduedate-property-outlook.md)|Returns or sets a  **Date** value that represents the due date of the task for this **[MailItem](mailitem-object-outlook.md)** . Read/write.|
|[TaskStartDate](mailitem-taskstartdate-property-outlook.md)|Returns or sets a  **Date** value that represents the start date of the task for this **[MailItem](mailitem-object-outlook.md)** object. Read/write.|
|[TaskSubject](mailitem-tasksubject-property-outlook.md)|Returns or sets a  **String** value that represents the subject of the task for the **[MailItem](mailitem-object-outlook.md)** object. Read/write.|
|[To](mailitem-to-property-outlook.md)|Returns or sets a semicolon-delimited  **String** list of display names for the **To** recipients for the Outlook item. Read/write.|
|[ToDoTaskOrdinal](mailitem-todotaskordinal-property-outlook.md)|Returns or sets a  **Date** value that represents the ordinal value of the task for the **[MailItem](mailitem-object-outlook.md)** . Read/write.|
|[UnRead](mailitem-unread-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** if the Outlook item has not been opened (read). Read/write.|
|[UserProperties](mailitem-userproperties-property-outlook.md)|Returns the  **[UserProperties](userproperties-object-outlook.md)** collection that represents all the user properties for the Outlook item. Read-only.|
|[VotingOptions](mailitem-votingoptions-property-outlook.md)|Returns or sets a  **String** specifying a delimited string containing the voting options for the mail message. Read/write.|
|[VotingResponse](mailitem-votingresponse-property-outlook.md)|Returns or sets a  **String** specifying the voting response for the mail message. Read/write.|

