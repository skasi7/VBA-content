---
title: SharingItem Members (Outlook)
ms.prod: OUTLOOK
ms.assetid: 719ad60e-2242-2c54-778f-006b61690389
---


# SharingItem Members (Outlook)
Represents a sharing message in an Inbox folder.

Represents a sharing message in an Inbox folder.


## Events



|**Name**|**Description**|
|:-----|:-----|
|[AfterWrite](sharingitem-afterwrite-event-outlook.md)|Occurs after Microsoft Outlook has saved the item.|
|[AttachmentAdd](sharingitem-attachmentadd-event-outlook.md)|Occurs when an attachment has been added to an instance of the parent object.|
|[AttachmentRead](sharingitem-attachmentread-event-outlook.md)|Occurs when an attachment in an instance of the parent object has been opened for reading.|
|[AttachmentRemove](sharingitem-attachmentremove-event-outlook.md)|Occurs when an attachment has been removed from an instance of the parent object.|
|[BeforeAttachmentAdd](sharingitem-beforeattachmentadd-event-outlook.md)|Occurs before an attachment is added to an instance of the parent object.|
|[BeforeAttachmentPreview](sharingitem-beforeattachmentpreview-event-outlook.md)|Occurs before an attachment associated with an instance of the parent object is previewed.|
|[BeforeAttachmentRead](sharingitem-beforeattachmentread-event-outlook.md)|Occurs before an attachment associated with an instance of the parent object is read.|
|[BeforeAttachmentSave](sharingitem-beforeattachmentsave-event-outlook.md)|Occurs before an attachment associated with an instance of the parent object is read.|
|[BeforeAttachmentWriteToTempFile](sharingitem-beforeattachmentwritetotempfile-event-outlook.md)|Occurs before an attachment associated with an instance of the parent object is written to a temporary file.|
|[BeforeAutoSave](sharingitem-beforeautosave-event-outlook.md)|Occurs before the  **[SharingItem](sharingitem-object-outlook.md)** is automatically saved by Outlook.|
|[BeforeCheckNames](sharingitem-beforechecknames-event-outlook.md)|Occurs just before Microsoft Outlook starts resolving names in the recipient collection for an item (which is an instance of the parent object).|
|[BeforeDelete](sharingitem-beforedelete-event-outlook.md)|Occurs before an item (which is an instance of the parent object) is deleted.|
|[BeforeRead](sharingitem-beforeread-event-outlook.md)|Occurs before Microsoft Outlook begins to read the properties for the item.|
|[Close](sharingitem-close-event-outlook.md)|Occurs when the inspector associated with an item (which is an instance of the parent object) is being closed.|
|[CustomAction](sharingitem-customaction-event-outlook.md)|Occurs when a custom action of an item (which is an instance of the parent object) executes.|
|[CustomPropertyChange](sharingitem-custompropertychange-event-outlook.md)|Occurs when a custom property of an item (which is an instance of the parent object) is changed. |
|[Forward](sharingitem-forward-event-outlook.md)|Occurs when the user selects the  **Forward** action for an item, or when the **[Forward](sharingitem-forward-method-outlook.md)** method is called for the item, which is an instance of the parent object.|
|[Open](sharingitem-open-event-outlook.md)|Occurs when an instance of the parent object is being opened in an  **[Inspector](inspector-object-outlook.md)** .|
|[PropertyChange](sharingitem-propertychange-event-outlook.md)|Occurs when an explicit built-in property (for example,  **[Subject](sharingitem-subject-property-outlook.md)** ) of an instance of the parent object is changed.|
|[Read](sharingitem-read-event-outlook.md)|Occurs when an instance of the parent object is opened for editing by the user. |
|[ReadComplete](sharingitem-readcomplete-event-outlook.md)|Occurs when Outlook has completed reading the properties of the item.|
|[Reply](sharingitem-reply-event-outlook.md)|Occurs when the user selects the  **Reply** action for an item, or when the **[Reply](sharingitem-reply-method-outlook.md)** method is called for the item, which is an instance of the parent object.|
|[ReplyAll](sharingitem-replyall-event-outlook.md)|Occurs when the user selects the  **ReplyAll** action for an item, or when the **[ReplyAll](sharingitem-replyall-method-outlook.md)** method is called for the item, which is an instance of the parent object.|
|[Send](sharingitem-send-event-outlook.md)|Occurs when the user selects the  **Send** action for an item, or when the **[Send](sharingitem-send-method-outlook.md)** method is called for the item, which is an instance of the parent object.|
|[Unload](sharingitem-unload-event-outlook.md)|Occurs before an Outlook item is unloaded from memory, either programmatically or by user action. |
|[Write](sharingitem-write-event-outlook.md)|Occurs when an instance of the parent object is saved, either explicitly (for example, using the  **[Save](sharingitem-save-method-outlook.md)** or **[SaveAs](sharingitem-saveas-method-outlook.md)** methods) or implicitly (for example, in response to a prompt when closing the item's inspector).|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[AddBusinessCard](sharingitem-addbusinesscard-method-outlook.md)|Appends contact information based on the Electronic Business Card (EBC) associated with the specified  **[ContactItem](contactitem-object-outlook.md)** object to the **[SharingItem](sharingitem-object-outlook.md)** object.|
|[Allow](sharingitem-allow-method-outlook.md)|Allows a sharing request and sends a sharing response to the sender of the  **[SharingItem](sharingitem-object-outlook.md)** .|
|[ClearConversationIndex](sharingitem-clearconversationindex-method-outlook.md)|Clears the index of the conversation thread for the  **[SharingItem](sharingitem-object-outlook.md)** .|
|[ClearTaskFlag](sharingitem-cleartaskflag-method-outlook.md)|Clears the  **[SharingItem](sharingitem-object-outlook.md)** object as a task.|
|[Close](sharingitem-close-method-outlook.md)|Closes and optionally saves changes to the  **[SharingItem](sharingitem-object-outlook.md)** .|
|[Copy](sharingitem-copy-method-outlook.md)|Creates another instance of a  **[SharingItem](sharingitem-object-outlook.md)** .|
|[Delete](sharingitem-delete-method-outlook.md)|Removes a  **[SharingItem](sharingitem-object-outlook.md)** item from the folder that contains the item.|
|[Deny](sharingitem-deny-method-outlook.md)|Denies a sharing request and sends a sharing response to the sender of the  **[SharingItem](sharingitem-object-outlook.md)** .|
|[Display](sharingitem-display-method-outlook.md)|Displays a new  **[Inspector](inspector-object-outlook.md)** object for the **[SharingItem](sharingitem-object-outlook.md)** .|
|[Forward](sharingitem-forward-method-outlook.md)|Executes the  **Forward** action for an item and returns the resulting copy as a **[SharingItem](sharingitem-object-outlook.md)** object.|
|[GetConversation](sharingitem-getconversation-method-outlook.md)|Obtains a  **[Conversation](conversation-object-outlook.md)** object that represents the conversation to which this item belongs.|
|[MarkAsTask](sharingitem-markastask-method-outlook.md)|Marks a  **[SharingItem](sharingitem-object-outlook.md)** object as a task and assigns a task interval for the object.|
|[Move](sharingitem-move-method-outlook.md)|Moves a  **[SharingItem](sharingitem-object-outlook.md)** to a new folder.|
|[OpenSharedFolder](sharingitem-opensharedfolder-method-outlook.md)|Opens a shared folder offered by a sharing invitation.|
|[PrintOut](sharingitem-printout-method-outlook.md)|Prints the  **[SharingItem](sharingitem-object-outlook.md)** using all default settings.|
|[Reply](sharingitem-reply-method-outlook.md)|Creates a reply, pre-addressed to the original sender, from the original  **[SharingItem](sharingitem-object-outlook.md)** .|
|[ReplyAll](sharingitem-replyall-method-outlook.md)|Creates a reply to all original recipients from the original  **[SharingItem](sharingitem-object-outlook.md)** .|
|[Save](sharingitem-save-method-outlook.md)|Saves the  **[SharingItem](sharingitem-object-outlook.md)** to the current folder or, if this is a new item, to the Outlook default folder for the item type.|
|[SaveAs](sharingitem-saveas-method-outlook.md)|Saves the  **[SharingItem](sharingitem-object-outlook.md)** to the specified path and in the format of the specified file type. If the file type is not specified, the MSG format (.msg) is used.|
|[Send](sharingitem-send-method-outlook.md)|Sends the  **[SharingItem](sharingitem-object-outlook.md)** .|
|[ShowCategoriesDialog](sharingitem-showcategoriesdialog-method-outlook.md)|Displays the  **Show Categories** dialog box, which allows you to select categories that correspond to the subject of the item.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Actions](sharingitem-actions-property-outlook.md)|Returns an  **[Actions](actions-object-outlook.md)** collection that represents all the available actions for the **[SharingItem](sharingitem-object-outlook.md)** . Read-only.|
|[AllowWriteAccess](sharingitem-allowwriteaccess-property-outlook.md)|Returns or sets a  **Boolean** value that indicates whether a sharing invitation should include write access to the folder. Read/write.|
|[AlternateRecipientAllowed](sharingitem-alternaterecipientallowed-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** if the item can be forwarded. Read/write.|
|[Application](sharingitem-application-property-outlook.md)|Returns an  **[Application](application-object-outlook.md)** object that represents the parent Outlook application for the **[SharingItem](sharingitem-object-outlook.md)** . Read-only.|
|[Attachments](sharingitem-attachments-property-outlook.md)|Returns an  **[Attachments](attachments-object-outlook.md)** object that represents all the attachments for the specified **[SharingItem](sharingitem-object-outlook.md)** . Read-only.|
|[AutoForwarded](sharingitem-autoforwarded-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** if the item was automatically forwarded. Read/write.|
|[BCC](sharingitem-bcc-property-outlook.md)|Returns a  **String** representing the display list of blind carbon copy (BCC) names for a **[SharingItem](sharingitem-object-outlook.md)** . Read/write.|
|[BillingInformation](sharingitem-billinginformation-property-outlook.md)|Returns or sets a  **String** representing the billing information associated with the **[SharingItem](sharingitem-object-outlook.md)** . Read/write.|
|[Body](sharingitem-body-property-outlook.md)|Returns or sets a  **String** representing the clear-text body of the **[SharingItem](sharingitem-object-outlook.md)** . Read/write.|
|[BodyFormat](sharingitem-bodyformat-property-outlook.md)|Returns or sets an  **[OlBodyFormat](olbodyformat-enumeration-outlook.md)** constant indicating the format of the body text. Read/write.|
|[Categories](sharingitem-categories-property-outlook.md)|Returns or sets a  **String** representing the categories assigned to the **[SharingItem](sharingitem-object-outlook.md)** . Read/write.|
|[CC](sharingitem-cc-property-outlook.md)|Returns a  **String** representing the display list of carbon copy (CC) names for a **[SharingItem](sharingitem-object-outlook.md)** . Read/write.|
|[Class](sharingitem-class-property-outlook.md)|Returns an  **[OlObjectClass](olobjectclass-enumeration-outlook.md)** constant indicating the object's class. Read-only.|
|[Companies](sharingitem-companies-property-outlook.md)|Returns or sets a  **String** representing the names of the companies associated with the **[SharingItem](sharingitem-object-outlook.md)** . Read/write.|
|[Conflicts](sharingitem-conflicts-property-outlook.md)|Return the  **[Conflicts](conflicts-object-outlook.md)** object that represents the items that are in conflict with the **[SharingItem](sharingitem-object-outlook.md)** . Read-only.|
|[ConversationID](sharingitem-conversationid-property-outlook.md)|Returns a  **String** that uniquely identifies a **[Conversation](conversation-object-outlook.md)** object that the **[SharingItem](sharingitem-object-outlook.md)** object belongs to. Read-only.|
|[ConversationIndex](sharingitem-conversationindex-property-outlook.md)|Returns a  **String** representing the index of the conversation thread of the **[SharingItem](sharingitem-object-outlook.md)** . Read-only.|
|[ConversationTopic](sharingitem-conversationtopic-property-outlook.md)|Returns a  **String** representing the topic of the conversation thread of the **[SharingItem](sharingitem-object-outlook.md)** . Read-only.|
|[CreationTime](sharingitem-creationtime-property-outlook.md)|Returns a  **Date** indicating the creation time for the **[SharingItem](sharingitem-object-outlook.md)** . Read-only.|
|[DeferredDeliveryTime](sharingitem-deferreddeliverytime-property-outlook.md)|Returns or sets a  **Date** indicating the date and time the **[SharingItem](sharingitem-object-outlook.md)** is to be delivered. Read/write.|
|[DeleteAfterSubmit](sharingitem-deleteaftersubmit-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** if a copy of the item is not saved upon being sent, and **False** if a copy is saved. Read/write.|
|[DownloadState](sharingitem-downloadstate-property-outlook.md)|Returns a constant that belongs to the  **[OlDownloadState](oldownloadstate-enumeration-outlook.md)** enumeration indicating the download state of the **[SharingItem](sharingitem-object-outlook.md)** . Read-only.|
|[EntryID](sharingitem-entryid-property-outlook.md)|Returns a  **String** representing the unique Entry ID of the **[SharingItem](sharingitem-object-outlook.md)** . Read-only.|
|[ExpiryTime](sharingitem-expirytime-property-outlook.md)|Returns or sets a  **Date** indicating the date and time at which the **[SharingItem](sharingitem-object-outlook.md)** becomes invalid and can be deleted. Read/write.|
|[FlagRequest](sharingitem-flagrequest-property-outlook.md)|Returns or sets a  **String** indicating the requested action for the **[SharingItem](sharingitem-object-outlook.md)** . Read/write.|
|[FormDescription](sharingitem-formdescription-property-outlook.md)|Returns the  **[FormDescription](formdescription-object-outlook.md)** object that represents the form description for the specified **[SharingItem](sharingitem-object-outlook.md)** . Read-only.|
|[GetInspector](sharingitem-getinspector-property-outlook.md)|Returns an  **[Inspector](inspector-object-outlook.md)** object that represents an inspector initialized to contain the specified **[SharingItem](sharingitem-object-outlook.md)** . Read-only.|
|[HTMLBody](sharingitem-htmlbody-property-outlook.md)|Returns or sets a  **String** representing the HTML body of the specified **[SharingItem](sharingitem-object-outlook.md)** . Read/write.|
|[Importance](sharingitem-importance-property-outlook.md)|Returns or sets an  **[OlImportance](olimportance-enumeration-outlook.md)** constant indicating the relative importance level for the **[SharingItem](sharingitem-object-outlook.md)** . Read/write.|
|[InternetCodepage](sharingitem-internetcodepage-property-outlook.md)|Returns or sets a  **Long** that determines the Internet code page used by the **[SharingItem](sharingitem-object-outlook.md)** . Read/write.|
|[IsConflict](sharingitem-isconflict-property-outlook.md)|Returns a  **Boolean** that determines if the **[SharingItem](sharingitem-object-outlook.md)** is in conflict. Read-only.|
|[IsMarkedAsTask](sharingitem-ismarkedastask-property-outlook.md)|Returns a  **Boolean** value that indicates whether the **[SharingItem](sharingitem-object-outlook.md)** is marked as a task. Read-only.|
|[ItemProperties](sharingitem-itemproperties-property-outlook.md)|Returns an  **[ItemProperties](itemproperties-object-outlook.md)** collection that represents all standard and user-defined properties associated with the **[SharingItem](sharingitem-object-outlook.md)** . Read-only.|
|[LastModificationTime](sharingitem-lastmodificationtime-property-outlook.md)|Returns a  **Date** specifying the date and time that the **[SharingItem](sharingitem-object-outlook.md)** was last modified. Read-only.|
|[MarkForDownload](sharingitem-markfordownload-property-outlook.md)|Returns or sets an  **[OlRemoteStatus](olremotestatus-enumeration-outlook.md)** constant that determines the status of an item once it is received by a remote user. Read/write.|
|[MessageClass](sharingitem-messageclass-property-outlook.md)|Returns or sets a  **String** representing the message class for the **[SharingItem](sharingitem-object-outlook.md)** . Read/write.|
|[Mileage](sharingitem-mileage-property-outlook.md)|Returns or sets a  **String** representing the mileage for a **[SharingItem](sharingitem-object-outlook.md)** . Read/write.|
|[NoAging](sharingitem-noaging-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** to not age the **[SharingItem](sharingitem-object-outlook.md)** . Read/write.|
|[OriginatorDeliveryReportRequested](sharingitem-originatordeliveryreportrequested-property-outlook.md)|Returns or sets a  **Boolean** value that determines whether the originator of the **[SharingItem](sharingitem-object-outlook.md)** will receive a delivery report. Read/write.|
|[OutlookInternalVersion](sharingitem-outlookinternalversion-property-outlook.md)|Returns a  **Long** representing the build number of the Outlook application for a **[SharingItem](sharingitem-object-outlook.md)** . Read-only.|
|[OutlookVersion](sharingitem-outlookversion-property-outlook.md)|Returns a  **String** indicating the major and minor version number of the Outlook application for a **[SharingItem](sharingitem-object-outlook.md)** . Read-only.|
|[Parent](sharingitem-parent-property-outlook.md)|Returns the parent  **Object** of the specified **[SharingItem](sharingitem-object-outlook.md)** . Read-only.|
|[Permission](sharingitem-permission-property-outlook.md)|Sets or returns an  **[OlPermission](olpermission-enumeration-outlook.md)** constant that determines what permissions to grant the recipients on the **[SharingItem](sharingitem-object-outlook.md)** . Read/write.|
|[PermissionService](sharingitem-permissionservice-property-outlook.md)|Sets or returns an  **[OlPermissionService](olpermissionservice-enumeration-outlook.md)** constant that determines the permission service that will be used when sending a **[SharingItem](sharingitem-object-outlook.md)** protected by Information Rights Management (IRM). Read/write.|
|[PermissionTemplateGuid](sharingitem-permissiontemplateguid-property-outlook.md)|Returns or sets a  **String** that represents the GUID of the template file to be applied to the **[SharingItem](sharingitem-object-outlook.md)** in order to specify Information Rights Management (IRM) permissions. Read/write.|
|[PropertyAccessor](sharingitem-propertyaccessor-property-outlook.md)|Returns a  **[PropertyAccessor](propertyaccessor-object-outlook.md)** object that supports creating, getting, setting, and deleting properties of the parent **[SharingItem](sharingitem-object-outlook.md)** object. Read-only.|
|[ReadReceiptRequested](sharingitem-readreceiptrequested-property-outlook.md)|Returns a  **Boolean** value that indicates **true** if a read receipt has been requested by the sender.|
|[ReceivedByEntryID](sharingitem-receivedbyentryid-property-outlook.md)|Returns a  **String** representing the **[EntryID](recipient-entryid-property-outlook.md)** for the true recipient as set by the transport provider delivering the **[SharingItem](sharingitem-object-outlook.md)** . Read-only.|
|[ReceivedByName](sharingitem-receivedbyname-property-outlook.md)|Returns a  **String** representing the display name of the true recipient for the **[SharingItem](sharingitem-object-outlook.md)** . Read-only.|
|[ReceivedOnBehalfOfEntryID](sharingitem-receivedonbehalfofentryid-property-outlook.md)|Returns a  **String** representing the **[EntryID](recipient-entryid-property-outlook.md)** of the user delegated to represent the recipient for the **[SharingItem](sharingitem-object-outlook.md)** . Read-only.|
|[ReceivedOnBehalfOfName](sharingitem-receivedonbehalfofname-property-outlook.md)|Returns a  **String** representing the display name of the user delegated to represent the recipient for the **[SharingItem](sharingitem-object-outlook.md)** . Read-only.|
|[ReceivedTime](sharingitem-receivedtime-property-outlook.md)|Returns a  **Date** indicating the date and time at which the **[SharingItem](sharingitem-object-outlook.md)** was received. Read-only.|
|[RecipientReassignmentProhibited](sharingitem-recipientreassignmentprohibited-property-outlook.md)|Returns a  **Boolean** that indicates **true** if the recipient cannot forward the specified **[SharingItem](sharingitem-object-outlook.md)** . Read/write.|
|[Recipients](sharingitem-recipients-property-outlook.md)|Returns a  **[Recipients](recipients-object-outlook.md)** collection that represents all the recipients for the **[SharingItem](sharingitem-object-outlook.md)** . Read-only.|
|[ReminderOverrideDefault](sharingitem-reminderoverridedefault-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** if the reminder overrides the default reminder behavior for the specified **[SharingItem](sharingitem-object-outlook.md)** . Read/write.|
|[ReminderPlaySound](sharingitem-reminderplaysound-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** if the reminder should play a sound when it occurs for the specified **[SharingItem](sharingitem-object-outlook.md)** . Read/write.|
|[ReminderSet](sharingitem-reminderset-property-outlook.md)|Returns or sets a  **Boolean** value that is **true** if a reminder has been set for the specified **[SharingItem](sharingitem-object-outlook.md)** . Read/write.|
|[ReminderSoundFile](sharingitem-remindersoundfile-property-outlook.md)|Returns or sets a  **String** indicating the path and file name of the sound file to play when the reminder occurs for the Outlook item. Read/write.|
|[ReminderTime](sharingitem-remindertime-property-outlook.md)|Returns or sets a  **Date** indicating the date and time at which the reminder should occur for the specified **[SharingItem](sharingitem-object-outlook.md)** . Read/write.|
|[RemoteID](sharingitem-remoteid-property-outlook.md)|Returns a  **String** that represents the unique identifier of the sharing context for a **[SharingItem](sharingitem-object-outlook.md)** object. Read-only.|
|[RemoteName](sharingitem-remotename-property-outlook.md)|Returns a  **String** that represents the name of the sharing context for a **[SharingItem](sharingitem-object-outlook.md)** object. Read-only.|
|[RemotePath](sharingitem-remotepath-property-outlook.md)|Returns a  **String** that represents the path of the sharing context for a **[SharingItem](sharingitem-object-outlook.md)** object. Read-only.|
|[RemoteStatus](sharingitem-remotestatus-property-outlook.md)|Returns or sets an  **[OlRemoteStatus](olremotestatus-enumeration-outlook.md)** constant specifying the remote status of the **[SharingItem](sharingitem-object-outlook.md)** . Read/write.|
|[ReplyRecipientNames](sharingitem-replyrecipientnames-property-outlook.md)|Returns a semicolon-delimited  **String** list of reply recipients for the **[SharingItem](sharingitem-object-outlook.md)** . Read-only.|
|[ReplyRecipients](sharingitem-replyrecipients-property-outlook.md)|Returns a  **[Recipients](recipients-object-outlook.md)** collection that represents all the reply recipient objects for the **[SharingItem](sharingitem-object-outlook.md)** . Read-only.|
|[RequestedFolder](sharingitem-requestedfolder-property-outlook.md)|Returns an  **[OlDefaultFolders](oldefaultfolders-enumeration-outlook.md)** constant that represents the type of default folder to which access is requested by a sharing request. Read-only.|
|[RetentionExpirationDate](sharingitem-retentionexpirationdate-property-outlook.md)|Returns a  **Date** that specifies the date when the **[SharingItem](sharingitem-object-outlook.md)** object expires, after which the Messaging Records Management (MRM) Assistant will delete the item. Read-only.|
|[RetentionPolicyName](sharingitem-retentionpolicyname-property-outlook.md)|Returns a  **String** that specifies the name of the retention policy. Read-only.|
|[RTFBody](sharingitem-rtfbody-property-outlook.md)|Returns or sets a  **Byte** array that represents the body of the Microsoft Outlook item in Rich Text Format. Read/write.|
|[Saved](sharingitem-saved-property-outlook.md)|Returns a  **Boolean** value that is **true** if the **[SharingItem](sharingitem-object-outlook.md)** has not been modified since the last save. Read-only.|
|[SaveSentMessageFolder](sharingitem-savesentmessagefolder-property-outlook.md)|Returns or sets a  **[Folder](folder-object-outlook.md)** object that represents the folder in which a copy of the **[SharingItem](sharingitem-object-outlook.md)** will be saved after being sent. Read/write.|
|[SenderEmailAddress](sharingitem-senderemailaddress-property-outlook.md)|Returns a  **String** that represents the e-mail address of the sender of the **[SharingItem](sharingitem-object-outlook.md)** . Read-only.|
|[SenderEmailType](sharingitem-senderemailtype-property-outlook.md)|Returns a  **String** that represents the type of entry for the e-mail address of the sender of the **[SharingItem](sharingitem-object-outlook.md)** , such as 'SMTP' for Internet address, 'EX' for a Microsoft Exchange server address, and so on. Read-only.|
|[SenderName](sharingitem-sendername-property-outlook.md)|Returns a  **String** indicating the display name of the sender for the **[SharingItem](sharingitem-object-outlook.md)** . Read-only.|
|[SendUsingAccount](sharingitem-sendusingaccount-property-outlook.md)|Returns or sets an  **[Account](account-object-outlook.md)** object that represents the account under which the **[SharingItem](sharingitem-object-outlook.md)** is to be sent. Read/write.|
|[Sensitivity](sharingitem-sensitivity-property-outlook.md)|Returns or sets an  **[OlSensitivity](olsensitivity-enumeration-outlook.md)** constant indicating the sensitivity for the **[SharingItem](sharingitem-object-outlook.md)** . Read/write.|
|[Sent](sharingitem-sent-property-outlook.md)|Returns a  **Boolean** value that indicates if the **[SharingItem](sharingitem-object-outlook.md)** has been sent. Read-only.|
|[SentOn](sharingitem-senton-property-outlook.md)|Returns a  **Date** indicating the date and time on which the **[SharingItem](sharingitem-object-outlook.md)** was sent. Read-only.|
|[SentOnBehalfOfName](sharingitem-sentonbehalfofname-property-outlook.md)|Returns or sets a  **String** indicating the display name for the intended sender of the **[SharingItem](sharingitem-object-outlook.md)** . Read/write.|
|[Session](sharingitem-session-property-outlook.md)|Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.|
|[SharingProvider](sharingitem-sharingprovider-property-outlook.md)|Returns an  **[OlSharingProvider](olsharingprovider-enumeration-outlook.md)** constant that indicates the sharing provider used by the **[SharingItem](sharingitem-object-outlook.md)** . Read-only.|
|[SharingProviderGuid](sharingitem-sharingproviderguid-property-outlook.md)|Returns a  **String** that represents the GUID of the sharing provider used by the **[SharingItem](sharingitem-object-outlook.md)** object. Read-only.|
|[Size](sharingitem-size-property-outlook.md)|Returns a  **Long** indicating the size (in bytes) of the **[SharingItem](sharingitem-object-outlook.md)** . Read-only.|
|[Subject](sharingitem-subject-property-outlook.md)|Returns or sets a  **String** indicating the subject for the **[SharingItem](sharingitem-object-outlook.md)** . Read/write.|
|[Submitted](sharingitem-submitted-property-outlook.md)|Returns a  **Boolean** value that is **True** if the **[SharingItem](sharingitem-object-outlook.md)** has been submitted. Read-only.|
|[TaskCompletedDate](sharingitem-taskcompleteddate-property-outlook.md)|Returns or sets a  **Date** value that represents the completion date of the task for this **[SharingItem](sharingitem-object-outlook.md)** . Read/write.|
|[TaskDueDate](sharingitem-taskduedate-property-outlook.md)|Returns or sets a  **Date** value that represents the due date of the task for this **[SharingItem](sharingitem-object-outlook.md)** . Read/write.|
|[TaskStartDate](sharingitem-taskstartdate-property-outlook.md)|Returns or sets a  **Date** value that represents the start date of the task for this **[SharingItem](sharingitem-object-outlook.md)** object. Read/write.|
|[TaskSubject](sharingitem-tasksubject-property-outlook.md)|Returns or sets a  **String** value that represents the subject of the task for the **[SharingItem](sharingitem-object-outlook.md)** object. Read/write.|
|[To](sharingitem-to-property-outlook.md)|Returns or sets a semicolon-delimited  **String** list of display names for the **To** recipients for the **[SharingItem](sharingitem-object-outlook.md)** . Read/write.|
|[ToDoTaskOrdinal](sharingitem-todotaskordinal-property-outlook.md)|Returns or sets a  **Date** value that represents the ordinal value of the task for the **[SharingItem](sharingitem-object-outlook.md)** . Read/write.|
|[Type](sharingitem-type-property-outlook.md)|Returns or sets an  **[OlSharingMsgType](olsharingmsgtype-enumeration-outlook.md)** constant that indicates the type of sharing message represented by the **[SharingItem](sharingitem-object-outlook.md)** . Read/write.|
|[UnRead](sharingitem-unread-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** if the **[SharingItem](sharingitem-object-outlook.md)** has not been opened (read). Read/write.|
|[UserProperties](sharingitem-userproperties-property-outlook.md)|Returns the  **[UserProperties](userproperties-object-outlook.md)** collection that represents all the user properties for the **[SharingItem](sharingitem-object-outlook.md)** . Read-only.|

