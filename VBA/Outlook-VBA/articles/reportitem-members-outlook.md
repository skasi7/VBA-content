---
title: ReportItem Members (Outlook)
ms.prod: OUTLOOK
ms.assetid: 5a5662dd-e969-bbd5-129b-44609ba1cf9f
---


# ReportItem Members (Outlook)
Represents a mail-delivery report in an Inbox folder. 

Represents a mail-delivery report in an Inbox folder. 


## Events



|**Name**|**Description**|
|:-----|:-----|
|[AfterWrite](reportitem-afterwrite-event-outlook.md)|Occurs after Microsoft Outlook has saved the item.|
|[AttachmentAdd](reportitem-attachmentadd-event-outlook.md)|Occurs when an attachment has been added to an instance of the parent object.|
|[AttachmentRead](reportitem-attachmentread-event-outlook.md)|Occurs when an attachment in an instance of the parent object has been opened for reading.|
|[AttachmentRemove](reportitem-attachmentremove-event-outlook.md)|Occurs when an attachment has been removed from an instance of the parent object.|
|[BeforeAttachmentAdd](reportitem-beforeattachmentadd-event-outlook.md)|Occurs before an attachment is added to an instance of the parent object.|
|[BeforeAttachmentPreview](reportitem-beforeattachmentpreview-event-outlook.md)|Occurs before an attachment associated with an instance of the parent object is previewed.|
|[BeforeAttachmentRead](reportitem-beforeattachmentread-event-outlook.md)|Occurs before an attachment associated with an instance of the parent object is read from the file system, an attachment stream, or an  **[Attachment](attachment-object-outlook.md)** object.|
|[BeforeAttachmentSave](reportitem-beforeattachmentsave-event-outlook.md)|Occurs just before an attachment is saved.|
|[BeforeAttachmentWriteToTempFile](reportitem-beforeattachmentwritetotempfile-event-outlook.md)|Occurs before an attachment associated with an instance of the parent object is written to a temporary file.|
|[BeforeAutoSave](reportitem-beforeautosave-event-outlook.md)|Occurs before the item is automatically saved by Outlook.|
|[BeforeCheckNames](reportitem-beforechecknames-event-outlook.md)|Occurs just before Microsoft Outlook starts resolving names in the recipient collection for an item (which is an instance of the parent object).|
|[BeforeDelete](reportitem-beforedelete-event-outlook.md)|Occurs before an item (which is an instance of the parent object) is deleted.|
|[BeforeRead](reportitem-beforeread-event-outlook.md)|Occurs before Microsoft Outlook begins to read the properties for the item.|
|[Close](reportitem-close-event-outlook.md)|Occurs when the inspector associated with an item (which is an instance of the parent object) is being closed.|
|[CustomAction](reportitem-customaction-event-outlook.md)|Occurs when a custom action of an item (which is an instance of the parent object) executes.|
|[CustomPropertyChange](reportitem-custompropertychange-event-outlook.md)|Occurs when a custom property of an item (which is an instance of the parent object) is changed. |
|[Forward](reportitem-forward-event-outlook.md)|Occurs when the user selects the  **Forward** action for an item (which is an instance of the parent object).|
|[Open](reportitem-open-event-outlook.md)|Occurs when an instance of the parent object is being opened in an  **[Inspector](inspector-object-outlook.md)** .|
|[PropertyChange](reportitem-propertychange-event-outlook.md)|Occurs when an explicit built-in property (for example,  **[Subject](appointmentitem-subject-property-outlook.md)** ) of an instance of the parent object is changed.|
|[Read](reportitem-read-event-outlook.md)|Occurs when an instance of the parent object is opened for editing by the user. |
|[ReadComplete](reportitem-readcomplete-event-outlook.md)|Occurs when Outlook has completed reading the properties of the item.|
|[Reply](reportitem-reply-event-outlook.md)|Occurs when the user selects the  **Reply** action for an item (which is an instance of the parent object).|
|[ReplyAll](reportitem-replyall-event-outlook.md)|Occurs when the user selects the  **ReplyAll** action for an item (which is an instance of the parent object).|
|[Send](reportitem-send-event-outlook.md)|Occurs when the user selects the  **Send** action for an item (which is an instance of the parent object).|
|[Unload](reportitem-unload-event-outlook.md)|Occurs before an Outlook item is unloaded from memory, either programmatically or by user action. |
|[Write](reportitem-write-event-outlook.md)|Occurs when an instance of the parent object is saved, either explicitly (for example, using the  **[Save](reportitem-save-method-outlook.md)** or **[SaveAs](reportitem-saveas-method-outlook.md)** methods) or implicitly (for example, in response to a prompt when closing the item's inspector).|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Close](reportitem-close-method-outlook.md)|Closes and optionally saves changes to the Outlook item.|
|[Copy](reportitem-copy-method-outlook.md)|Creates another instance of an object.|
|[Delete](reportitem-delete-method-outlook.md)|Removes the item from the folder that contains the item.|
|[Display](reportitem-display-method-outlook.md)|Displays a new  **[Inspector](inspector-object-outlook.md)** object for the item.|
|[GetConversation](reportitem-getconversation-method-outlook.md)|Obtains a  **[Conversation](conversation-object-outlook.md)** object that represents the conversation to which this item belongs.|
|[Move](reportitem-move-method-outlook.md)|Moves a Microsoft Outlook item to a new folder.|
|[PrintOut](reportitem-printout-method-outlook.md)|Prints the Outlook item using all default settings.The  **PrintOut** method is the only Outlook method that can be used for printing.|
|[Save](reportitem-save-method-outlook.md)|Saves the Microsoft Outlook item to the current folder or, if this is a new item, to the Outlook default folder for the item type.|
|[SaveAs](reportitem-saveas-method-outlook.md)|Saves the Microsoft Outlook item to the specified path and in the format of the specified file type. If the file type is not specified, the MSG format (.msg) is used.|
|[ShowCategoriesDialog](reportitem-showcategoriesdialog-method-outlook.md)|Displays the  **Show Categories** dialog box, which allows you to select categories that correspond to the subject of the item.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Actions](reportitem-actions-property-outlook.md)|Returns an  **[Actions](actions-object-outlook.md)** collection that represents all the available actions for the item. Read-only.|
|[Application](reportitem-application-property-outlook.md)|Returns an  **[Application](application-object-outlook.md)** object that represents the parent Outlook application for the object. Read-only.|
|[Attachments](reportitem-attachments-property-outlook.md)|Returns an  **[Attachments](attachments-object-outlook.md)** object that represents all the attachments for the specified item. Read-only.|
|[AutoResolvedWinner](reportitem-autoresolvedwinner-property-outlook.md)|Returns a  **Boolean** that determines if the item is a winner of an automatic conflict resolution. Read-only.|
|[BillingInformation](reportitem-billinginformation-property-outlook.md)|Returns or sets a  **String** representing the billing information associated with the Outlook item. Read/write.|
|[Body](reportitem-body-property-outlook.md)|Returns or sets a  **String** representing the clear-text body of the Outlook item. Read/write.|
|[Categories](reportitem-categories-property-outlook.md)|Returns or sets a  **String** representing the categories assigned to the Outlook item. Read/write.|
|[Class](reportitem-class-property-outlook.md)|Returns an  **[OlObjectClass](olobjectclass-enumeration-outlook.md)** constant indicating the object's class. Read-only.|
|[Companies](reportitem-companies-property-outlook.md)|Returns or sets a  **String** representing the names of the companies associated with the Outlook item. Read/write.|
|[Conflicts](reportitem-conflicts-property-outlook.md)|Return the  **[Conflicts](conflicts-object-outlook.md)** object that represents the items that are in conflict for any Outlook item object. Read-only.|
|[ConversationID](reportitem-conversationid-property-outlook.md)|Returns a  **String** that uniquely identifies a **[Conversation](conversation-object-outlook.md)** object that the **[ReportItem](reportitem-object-outlook.md)** object belongs to. Read-only.|
|[ConversationIndex](reportitem-conversationindex-property-outlook.md)|Returns a  **String** that indicates the relative position of the item within the conversation thread. Read-only.|
|[ConversationTopic](reportitem-conversationtopic-property-outlook.md)|Returns a  **String** representing the topic of the conversation thread of the Outlook item. Read-only.|
|[CreationTime](reportitem-creationtime-property-outlook.md)|Returns a  **Date** indicating the creation time for the Outlook item. Read-only.|
|[DownloadState](reportitem-downloadstate-property-outlook.md)|Returns a constant that belongs to the  **[OlDownloadState](oldownloadstate-enumeration-outlook.md)** enumeration indicating the download state of the item. Read-only.|
|[EntryID](reportitem-entryid-property-outlook.md)|Returns a  **String** representing the unique Entry ID of the object. Read-only.|
|[FormDescription](reportitem-formdescription-property-outlook.md)|Returns the  **[FormDescription](formdescription-object-outlook.md)** object that represents the form description for the specified Outlook item. Read-only.|
|[GetInspector](reportitem-getinspector-property-outlook.md)|Returns an  **[Inspector](inspector-object-outlook.md)** object that represents an inspector initialized to contain the specified item. Read-only.|
|[Importance](reportitem-importance-property-outlook.md)|Returns or sets an  **[OlImportance](olimportance-enumeration-outlook.md)** constant indicating the relative importance level for the Outlook item. Read/write.|
|[IsConflict](reportitem-isconflict-property-outlook.md)|Returns a  **Boolean** that determines if the item is in conflict. Read-only.|
|[ItemProperties](reportitem-itemproperties-property-outlook.md)|Returns an  **[ItemProperties](itemproperties-object-outlook.md)** collection that represents all standard and user-defined properties associated with the Outlook item. Read-only.|
|[LastModificationTime](reportitem-lastmodificationtime-property-outlook.md)|Returns a  **Date** specifying the date and time that the Outlook item was last modified. Read-only.|
|[MarkForDownload](reportitem-markfordownload-property-outlook.md)|Returns or sets an  **[OlRemoteStatus](olremotestatus-enumeration-outlook.md)** constant that determines the status of an item once it is received by a remote user. Read/write.|
|[MessageClass](reportitem-messageclass-property-outlook.md)|Returns or sets a  **String** representing the message class for the Outlook item. Read/write.|
|[Mileage](reportitem-mileage-property-outlook.md)|Returns or sets a  **String** representing the mileage for an item. Read/write.|
|[NoAging](reportitem-noaging-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** to not age the Outlook item. Read/write.|
|[OutlookInternalVersion](reportitem-outlookinternalversion-property-outlook.md)|Returns a  **Long** representing the build number of the Outlook application for an Outlook item. Read-only.|
|[OutlookVersion](reportitem-outlookversion-property-outlook.md)|Returns a  **String** indicating the major and minor version number of the Outlook application for an Outlook item. Read-only.|
|[Parent](reportitem-parent-property-outlook.md)|Returns the parent  **Object** of the specified object. Read-only.|
|[PropertyAccessor](reportitem-propertyaccessor-property-outlook.md)|Returns a  **[PropertyAccessor](propertyaccessor-object-outlook.md)** object that supports creating, getting, setting, and deleting properties of the parent **[ReportItem](reportitem-object-outlook.md)** object. Read-only.|
|[RetentionExpirationDate](reportitem-retentionexpirationdate-property-outlook.md)|Returns a  **Date** that specifies the date when the **[ReportItem](reportitem-object-outlook.md)** object expires, after which the Messaging Records Management (MRM) Assistant will delete the item. Read-only.|
|[RetentionPolicyName](reportitem-retentionpolicyname-property-outlook.md)|Returns a  **String** that specifies the name of the retention policy. Read-only.|
|[Saved](reportitem-saved-property-outlook.md)|Returns a  **Boolean** value that is **True** if the Outlook item has not been modified since the last save. Read-only.|
|[Sensitivity](reportitem-sensitivity-property-outlook.md)|Returns or sets a constant in the  **[OlSensitivity](olsensitivity-enumeration-outlook.md)** enumeration indicating the sensitivity for the Outlook item. Read/write.|
|[Session](reportitem-session-property-outlook.md)|Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.|
|[Size](reportitem-size-property-outlook.md)|Returns a  **Long** indicating the size (in bytes) of the Outlook item. Read-only.|
|[Subject](reportitem-subject-property-outlook.md)|Returns or sets a  **String** indicating the subject for the Outlook item. Read/write.|
|[UnRead](reportitem-unread-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** if the Outlook item has not been opened (read). Read/write.|
|[UserProperties](reportitem-userproperties-property-outlook.md)|Returns the  **[UserProperties](userproperties-object-outlook.md)** collection that represents all the user properties for the Outlook item. Read-only.|

