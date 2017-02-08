---
title: MeetingItem Members (Outlook)
ms.prod: OUTLOOK
ms.assetid: 9ae6a19d-d326-4c37-90d8-5ed9933672a0
---


# MeetingItem Members (Outlook)
Represents a change to the recipient's Calendar folder initiated by another party or as a result of a group action.

Represents a change to the recipient's Calendar folder initiated by another party or as a result of a group action.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[AfterWrite](meetingitem-afterwrite-event-outlook.md)|Occurs after Microsoft Outlook has saved the item.|
|[AttachmentAdd](meetingitem-attachmentadd-event-outlook.md)|Occurs when an attachment has been added to an instance of the parent object.|
|[AttachmentRead](meetingitem-attachmentread-event-outlook.md)|Occurs when an attachment in an instance of the parent object has been opened for reading.|
|[AttachmentRemove](meetingitem-attachmentremove-event-outlook.md)|Occurs when an attachment has been removed from an instance of the parent object.|
|[BeforeAttachmentAdd](meetingitem-beforeattachmentadd-event-outlook.md)|Occurs before an attachment is added to an instance of the parent object.|
|[BeforeAttachmentPreview](meetingitem-beforeattachmentpreview-event-outlook.md)|Occurs before an attachment associated with an instance of the parent object is previewed.|
|[BeforeAttachmentRead](meetingitem-beforeattachmentread-event-outlook.md)|Occurs before an attachment associated with an instance of the parent object is read from the file system, an attachment stream, or an  **[Attachment](attachment-object-outlook.md)** object.|
|[BeforeAttachmentSave](meetingitem-beforeattachmentsave-event-outlook.md)|Occurs just before an attachment is saved.|
|[BeforeAttachmentWriteToTempFile](meetingitem-beforeattachmentwritetotempfile-event-outlook.md)|Occurs before an attachment associated with an instance of the parent object is written to a temporary file.|
|[BeforeAutoSave](meetingitem-beforeautosave-event-outlook.md)|Occurs before the item is automatically saved by Outlook.|
|[BeforeCheckNames](meetingitem-beforechecknames-event-outlook.md)|Occurs just before Microsoft Outlook starts resolving names in the recipient collection for an item (which is an instance of the parent object).|
|[BeforeDelete](meetingitem-beforedelete-event-outlook.md)|Occurs before an item (which is an instance of the parent object) is deleted.|
|[BeforeRead](meetingitem-beforeread-event-outlook.md)|Occurs before Microsoft Outlook begins to read the properties for the item.|
|[Close](meetingitem-close-event-outlook.md)|Occurs when the inspector associated with an item (which is an instance of the parent object) is being closed.|
|[CustomAction](meetingitem-customaction-event-outlook.md)|Occurs when a custom action of an item (which is an instance of the parent object) executes.|
|[CustomPropertyChange](meetingitem-custompropertychange-event-outlook.md)|Occurs when a custom property of an item (which is an instance of the parent object) is changed. |
|[Forward](meetingitem-forward-event-outlook.md)|Occurs when the user selects the  **Forward** action for an item, or when the **Forward** method is called for the item, which is an instance of the parent object.|
|[Open](meetingitem-open-event-outlook.md)|Occurs when an instance of the parent object is being opened in an  **[Inspector](inspector-object-outlook.md)** .|
|[PropertyChange](meetingitem-propertychange-event-outlook.md)|Occurs when an explicit built-in property (for example,  **[Subject](appointmentitem-subject-property-outlook.md)** ) of an instance of the parent object is changed.|
|[Read](meetingitem-read-event-outlook.md)|Occurs when an instance of the parent object is opened for editing by the user. |
|[ReadComplete](meetingitem-readcomplete-event-outlook.md)|Occurs when Outlook has completed reading the properties of the item.|
|[Reply](meetingitem-reply-event-outlook.md)|Occurs when the user selects the  **Reply** action for an item, or when the **Reply** method is called for the item, which is an instance of the parent object.|
|[ReplyAll](meetingitem-replyall-event-outlook.md)|Occurs when the user selects the  **ReplyAll** action for an item, or when the **ReplyAll** method is called for the item, which is an instance of the parent object.|
|[Send](meetingitem-send-event-outlook.md)|Occurs when the user selects the  **Send** action for an item, or when the **Send** method is called for the item, which is an instance of the parent object.|
|[Unload](meetingitem-unload-event-outlook.md)|Occurs before an Outlook item is unloaded from memory, either programmatically or by user action. |
|[Write](meetingitem-write-event-outlook.md)|Occurs when an instance of the parent object is saved, either explicitly (for example, using the  **[Save](meetingitem-save-method-outlook.md)** or **[SaveAs](meetingitem-saveas-method-outlook.md)** methods) or implicitly (for example, in response to a prompt when closing the item's inspector).|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Close](meetingitem-close-method-outlook.md)|Closes and optionally saves changes to the Outlook item.|
|[Copy](meetingitem-copy-method-outlook.md)|Creates another instance of an object.|
|[Delete](meetingitem-delete-method-outlook.md)|Removes the item from the folder that contains the item.|
|[Display](meetingitem-display-method-outlook.md)|Displays a new  **[Inspector](inspector-object-outlook.md)** object for the item.|
|[Forward](meetingitem-forward-method-outlook.md)|Executes the  **Forward** action for an item and returns the resulting copy as a **[MeetingItem](meetingitem-object-outlook.md)** object.|
|[GetAssociatedAppointment](meetingitem-getassociatedappointment-method-outlook.md)|Returns an  **[AppointmentItem](appointmentitem-object-outlook.md)** object that represents the appointment associated with the meeting request.|
|[GetConversation](meetingitem-getconversation-method-outlook.md)|Obtains a  **[Conversation](conversation-object-outlook.md)** object that represents the conversation to which this item belongs.|
|[Move](meetingitem-move-method-outlook.md)|Moves a Microsoft Outlook item to a new folder.|
|[PrintOut](meetingitem-printout-method-outlook.md)|Prints the Outlook item using all default settings.The  **PrintOut** method is the only Outlook method that can be used for printing.|
|[Reply](meetingitem-reply-method-outlook.md)|Creates a reply, pre-addressed to the original sender, from the original message.|
|[ReplyAll](meetingitem-replyall-method-outlook.md)|Creates a reply to all original recipients from the original message.|
|[Save](meetingitem-save-method-outlook.md)|Saves the Microsoft Outlook item to the current folder or, if this is a new item, to the Outlook default folder for the item type.|
|[SaveAs](meetingitem-saveas-method-outlook.md)|Saves the Microsoft Outlook item to the specified path and in the format of the specified file type. If the file type is not specified, the MSG format (.msg) is used.|
|[Send](meetingitem-send-method-outlook.md)|Sends the meeting item.|
|[ShowCategoriesDialog](meetingitem-showcategoriesdialog-method-outlook.md)|Displays the  **Show Categories** dialog box, which allows you to select categories that correspond to the subject of the item.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Actions](meetingitem-actions-property-outlook.md)|Returns an  **[Actions](actions-object-outlook.md)** collection that represents all the available actions for the item. Read-only.|
|[Application](meetingitem-application-property-outlook.md)|Returns an  **[Application](application-object-outlook.md)** object that represents the parent Outlook application for the object. Read-only.|
|[Attachments](meetingitem-attachments-property-outlook.md)|Returns an  **[Attachments](attachments-object-outlook.md)** object that represents all the attachments for the specified item. Read-only.|
|[AutoForwarded](meetingitem-autoforwarded-property-outlook.md)|A  **Boolean** value that returns **True** if the item was automatically forwarded. Read/write.|
|[AutoResolvedWinner](meetingitem-autoresolvedwinner-property-outlook.md)|Returns a  **Boolean** that determines if the item is a winner of an automatic conflict resolution. Read-only.|
|[BillingInformation](meetingitem-billinginformation-property-outlook.md)|Returns or sets a  **String** representing the billing information associated with the Outlook item. Read/write.|
|[Body](meetingitem-body-property-outlook.md)|Returns or sets a  **String** representing the clear-text body of the Outlook item. Read/write.|
|[Categories](meetingitem-categories-property-outlook.md)|Returns or sets a  **String** representing the categories assigned to the Outlook item. Read/write.|
|[Class](meetingitem-class-property-outlook.md)|Returns an  **[OlObjectClass](olobjectclass-enumeration-outlook.md)** constant indicating the object's class. Read-only.|
|[Companies](meetingitem-companies-property-outlook.md)|Returns or sets a  **String** representing the names of the companies associated with the Outlook item. Read/write.|
|[Conflicts](meetingitem-conflicts-property-outlook.md)|Return the  **[Conflicts](conflicts-object-outlook.md)** object that represents the items that are in conflict for any Outlook item object. Read-only.|
|[ConversationID](meetingitem-conversationid-property-outlook.md)|Returns a  **String** that uniquely identifies a **[Conversation](conversation-object-outlook.md)** object that the **[MeetingItem](meetingitem-object-outlook.md)** object belongs to. Read-only.|
|[ConversationIndex](meetingitem-conversationindex-property-outlook.md)|Returns a  **String** that indicates the relative position of the item within the conversation thread. Read-only.|
|[ConversationTopic](meetingitem-conversationtopic-property-outlook.md)|Returns a  **String** representing the topic of the conversation thread of the Outlook item. Read-only.|
|[CreationTime](meetingitem-creationtime-property-outlook.md)|Returns a  **Date** indicating the creation time for the Outlook item. Read-only.|
|[DeferredDeliveryTime](meetingitem-deferreddeliverytime-property-outlook.md)|Returns or sets a  **Date** indicating the date and time the mail message is to be delivered. Read/write.|
|[DeleteAfterSubmit](meetingitem-deleteaftersubmit-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** if a copy of the mail message is not saved upon being sent, and **False** if a copy is saved. Read/write.|
|[DownloadState](meetingitem-downloadstate-property-outlook.md)|Returns a constant that belongs to the  **[OlDownloadState](oldownloadstate-enumeration-outlook.md)** enumeration indicating the download state of the item. Read-only.|
|[EntryID](meetingitem-entryid-property-outlook.md)|Returns a  **String** representing the unique Entry ID of the object. Read-only.|
|[ExpiryTime](meetingitem-expirytime-property-outlook.md)|Returns or sets a  **Date** indicating the date and time at which the item becomes invalid and can be deleted. Read/write.|
|[FormDescription](meetingitem-formdescription-property-outlook.md)|Returns the  **[FormDescription](formdescription-object-outlook.md)** object that represents the form description for the specified Outlook item. Read-only.|
|[GetInspector](meetingitem-getinspector-property-outlook.md)|Returns an  **[Inspector](inspector-object-outlook.md)** object that represents an inspector initialized to contain the specified item. Read-only.|
|[Importance](meetingitem-importance-property-outlook.md)|Returns or sets an  **[OlImportance](olimportance-enumeration-outlook.md)** constant indicating the relative importance level for the Outlook item. Read/write.|
|[IsConflict](meetingitem-isconflict-property-outlook.md)|Returns a  **Boolean** that determines if the item is in conflict. Read-only.|
|[IsLatestVersion](meetingitem-islatestversion-property-outlook.md)|Returns a  **Boolean** value that indicates whether the **[MeetingItem](meetingitem-object-outlook.md)** represents the latest version of the item on the organizer's calendar. Read-only.|
|[ItemProperties](meetingitem-itemproperties-property-outlook.md)|Returns an  **[ItemProperties](itemproperties-object-outlook.md)** collection that represents all standard and user-defined properties associated with the Outlook item. Read-only.|
|[LastModificationTime](meetingitem-lastmodificationtime-property-outlook.md)|Returns a  **Date** specifying the date and time that the Outlook item was last modified. Read-only.|
|[MarkForDownload](meetingitem-markfordownload-property-outlook.md)|Returns or sets an  **[OlRemoteStatus](olremotestatus-enumeration-outlook.md)** constant that determines the status of an item once it is received by a remote user. Read/write.|
|[MeetingWorkspaceURL](meetingitem-meetingworkspaceurl-property-outlook.md)|Returns a  **String** value that represents the URL for the Meeting Workspace that the meeting item is linked to. Read-only.|
|[MessageClass](meetingitem-messageclass-property-outlook.md)|Returns or sets a  **String** representing the message class for the Outlook item. Read/write.|
|[Mileage](meetingitem-mileage-property-outlook.md)|Returns or sets a  **String** representing the mileage for an item. Read/write.|
|[NoAging](meetingitem-noaging-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** to not age the Outlook item. Read/write.|
|[OriginatorDeliveryReportRequested](meetingitem-originatordeliveryreportrequested-property-outlook.md)|Returns or sets a  **Boolean** value that determines whether the originator of the meeting item or mail message will receive a delivery report. Read/write.|
|[OutlookInternalVersion](meetingitem-outlookinternalversion-property-outlook.md)|Returns a  **Long** representing the build number of the Outlook application for an Outlook item. Read-only.|
|[OutlookVersion](meetingitem-outlookversion-property-outlook.md)|Returns a  **String** indicating the major and minor version number of the Outlook application for an Outlook item. Read-only.|
|[Parent](meetingitem-parent-property-outlook.md)|Returns the parent  **Object** of the specified object. Read-only.|
|[PropertyAccessor](meetingitem-propertyaccessor-property-outlook.md)|Returns a  **[PropertyAccessor](propertyaccessor-object-outlook.md)** object that supports creating, getting, setting, and deleting properties of the parent **[MeetingItem](meetingitem-object-outlook.md)** object. Read-only.|
|[ReceivedTime](meetingitem-receivedtime-property-outlook.md)|Returns or sets a  **Date** indicating the date and time at which the item was received. Read/write.|
|[Recipients](meetingitem-recipients-property-outlook.md)|Returns a  **[Recipients](recipients-object-outlook.md)** collection that represents all the recipients for the Outlook item. Read-only.|
|[ReminderSet](meetingitem-reminderset-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** if a reminder has been set for this item. Read/write.|
|[ReminderTime](meetingitem-remindertime-property-outlook.md)|Returns or sets a  **Date** indicating the date and time at which the reminder should occur for the specified item. Read/write.|
|[ReplyRecipients](meetingitem-replyrecipients-property-outlook.md)|Returns a  **[Recipients](recipients-object-outlook.md)** collection that represents all the reply recipient objects for the Outlook item. Read-only.|
|[RetentionExpirationDate](meetingitem-retentionexpirationdate-property-outlook.md)|Returns a  **Date** that specifies the date when the **[MeetingItem](meetingitem-object-outlook.md)** object expires, after which the Messaging Records Management (MRM) Assistant will delete the item. Read-only.|
|[RetentionPolicyName](meetingitem-retentionpolicyname-property-outlook.md)|Returns a  **String** that specifies the name of the retention policy. Read-only.|
|[RTFBody](meetingitem-rtfbody-property-outlook.md)|Returns or sets a  **Byte** array that represents the body of the Microsoft Outlook item in Rich Text Format. Read/write.|
|[Saved](meetingitem-saved-property-outlook.md)|Returns a  **Boolean** value that is **True** if the Outlook item has not been modified since the last save. Read-only.|
|[SaveSentMessageFolder](meetingitem-savesentmessagefolder-property-outlook.md)|Setting or getting this property has no noticeable effect. Do not use this property.|
|[SenderEmailAddress](meetingitem-senderemailaddress-property-outlook.md)|Returns a  **String** that represents the e-mail address of the sender of the Outlook item. Read-only.|
|[SenderEmailType](meetingitem-senderemailtype-property-outlook.md)|Returns a  **String** that represents the type of entry for the e-mail address of the sender of the Outlook item, such as 'SMTP' for Internet address, 'EX' for a Microsoft Exchange server address, etc. Read-only.|
|[SenderName](meetingitem-sendername-property-outlook.md)|Returns a  **String** indicating the display name of the sender for the Outlook item. Read-only.|
|[SendUsingAccount](meetingitem-sendusingaccount-property-outlook.md)|Returns or sets an  **[Account](account-object-outlook.md)** object that represents the account to use to send the **[MeetingItem](meetingitem-object-outlook.md)** . Read/write.|
|[Sensitivity](meetingitem-sensitivity-property-outlook.md)|Returns or sets a constant in the  **[OlSensitivity](olsensitivity-enumeration-outlook.md)** enumeration indicating the sensitivity for the Outlook item. Read/write.|
|[Sent](meetingitem-sent-property-outlook.md)|Returns a  **Boolean** value that indicates if a message has been sent. Read-only.|
|[SentOn](meetingitem-senton-property-outlook.md)|Returns a  **Date** indicating the date and time on which the Outlook item was sent. Read-only.|
|[Session](meetingitem-session-property-outlook.md)|Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.|
|[Size](meetingitem-size-property-outlook.md)|Returns a  **Long** indicating the size (in bytes) of the Outlook item. Read-only.|
|[Subject](meetingitem-subject-property-outlook.md)|Returns or sets a  **String** indicating the subject for the Outlook item. Read/write.|
|[Submitted](meetingitem-submitted-property-outlook.md)|Returns a  **Boolean** value that is **True** if the item has been submitted. Read-only.|
|[UnRead](meetingitem-unread-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** if the Outlook item has not been opened (read). Read/write.|
|[UserProperties](meetingitem-userproperties-property-outlook.md)|Returns the  **[UserProperties](userproperties-object-outlook.md)** collection that represents all the user properties for the Outlook item. Read-only.|

