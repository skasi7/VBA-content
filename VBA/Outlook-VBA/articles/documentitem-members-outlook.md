---
title: DocumentItem Members (Outlook)
ms.prod: OUTLOOK
ms.assetid: 2c6d563b-39cb-9cb3-3bfe-93fe595325cf
---


# DocumentItem Members (Outlook)
Represents any document other than a Microsoft Outlook item as an item in an Outlook folder. 

Represents any document other than a Microsoft Outlook item as an item in an Outlook folder. 


## Events



|**Name**|**Description**|
|:-----|:-----|
|[AfterWrite](documentitem-afterwrite-event-outlook.md)|Occurs after Microsoft Outlook has saved the item.|
|[AttachmentAdd](documentitem-attachmentadd-event-outlook.md)|Occurs when an attachment has been added to an instance of the parent object.|
|[AttachmentRead](documentitem-attachmentread-event-outlook.md)|Occurs when an attachment in an instance of the parent object has been opened for reading.|
|[AttachmentRemove](documentitem-attachmentremove-event-outlook.md)|Occurs when an attachment has been removed from an instance of the parent object.|
|[BeforeAttachmentAdd](documentitem-beforeattachmentadd-event-outlook.md)|Occurs before an attachment is added to an instance of the parent object.|
|[BeforeAttachmentPreview](documentitem-beforeattachmentpreview-event-outlook.md)|Occurs before an attachment associated with an instance of the parent object is previewed.|
|[BeforeAttachmentRead](documentitem-beforeattachmentread-event-outlook.md)|Occurs before an attachment associated with an instance of the parent object is read from the file system, an attachment stream, or an  **[Attachment](attachment-object-outlook.md)** object.|
|[BeforeAttachmentSave](documentitem-beforeattachmentsave-event-outlook.md)|Occurs just before an attachment is saved.|
|[BeforeAttachmentWriteToTempFile](documentitem-beforeattachmentwritetotempfile-event-outlook.md)|Occurs before an attachment associated with an instance of the parent object is written to a temporary file.|
|[BeforeAutoSave](documentitem-beforeautosave-event-outlook.md)|Occurs before the item is automatically saved by Outlook.|
|[BeforeCheckNames](documentitem-beforechecknames-event-outlook.md)|Occurs just before Microsoft Outlook starts resolving names in the recipient collection for an item (which is an instance of the parent object).|
|[BeforeDelete](documentitem-beforedelete-event-outlook.md)|Occurs before an item (which is an instance of the parent object) is deleted.|
|[BeforeRead](documentitem-beforeread-event-outlook.md)|Occurs before Microsoft Outlook begins to read the properties for the item.|
|[Close](documentitem-close-event-outlook.md)|Occurs when the inspector associated with an item (which is an instance of the parent object) is being closed.|
|[CustomAction](documentitem-customaction-event-outlook.md)|Occurs when a custom action of an item (which is an instance of the parent object) executes.|
|[CustomPropertyChange](documentitem-custompropertychange-event-outlook.md)|Occurs when a custom property of an item (which is an instance of the parent object) is changed. |
|[Forward](documentitem-forward-event-outlook.md)|Occurs when the user selects the  **Forward** action for an item (which is an instance of the parent object).|
|[Open](documentitem-open-event-outlook.md)|Occurs when an instance of the parent object is being opened in an  **[Inspector](inspector-object-outlook.md)** .|
|[PropertyChange](documentitem-propertychange-event-outlook.md)|Occurs when an explicit built-in property (for example,  **[Subject](appointmentitem-subject-property-outlook.md)** ) of an instance of the parent object is changed.|
|[Read](documentitem-read-event-outlook.md)|Occurs when an instance of the parent object is opened for editing by the user. |
|[ReadComplete](documentitem-readcomplete-event-outlook.md)|Occurs when Outlook has completed reading the properties of the item.|
|[Reply](documentitem-reply-event-outlook.md)|Occurs when the user selects the  **Reply** action for an item (which is an instance of the parent object).|
|[ReplyAll](documentitem-replyall-event-outlook.md)|Occurs when the user selects the  **ReplyAll** action for an item (which is an instance of the parent object).|
|[Send](documentitem-send-event-outlook.md)|Occurs when the user selects the  **Send** action for an item (which is an instance of the parent object).|
|[Unload](documentitem-unload-event-outlook.md)|Occurs before an Outlook item is unloaded from memory, either programmatically or by user action. |
|[Write](documentitem-write-event-outlook.md)|Occurs when an instance of the parent object is saved, either explicitly (for example, using the  **[Save](documentitem-save-method-outlook.md)** or **[SaveAs](documentitem-saveas-method-outlook.md)** methods) or implicitly (for example, in response to a prompt when closing the item's inspector).|

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[Close](documentitem-close-method-outlook.md)|This object, member, or enumeration will be deprecated and is not intended to be used in your code. |
|[Copy](documentitem-copy-method-outlook.md)|Creates another instance of an object.|
|[Delete](documentitem-delete-method-outlook.md)|Removes the item from the folder that contains the item.|
|[Display](documentitem-display-method-outlook.md)|Displays a new  **[Inspector](inspector-object-outlook.md)** object for the item.|
|[Move](documentitem-move-method-outlook.md)|Moves a Microsoft Outlook item to a new folder.|
|[PrintOut](documentitem-printout-method-outlook.md)|Prints the Outlook item using all default settings.The  **PrintOut** method is the only Outlook method that can be used for printing.|
|[Save](documentitem-save-method-outlook.md)|Saves the Microsoft Outlook item to the current folder or, if this is a new item, to the Outlook default folder for the item type.|
|[SaveAs](documentitem-saveas-method-outlook.md)|Saves the Microsoft Outlook item to the specified path and in the format of the specified file type. If the file type is not specified, the MSG format (.msg) is used.|
|[ShowCategoriesDialog](documentitem-showcategoriesdialog-method-outlook.md)|Displays the  **Show Categories** dialog box, which allows you to select categories that correspond to the subject of the item.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Actions](documentitem-actions-property-outlook.md)|Returns an  **[Actions](actions-object-outlook.md)** collection that represents all the available actions for the item. Read-only.|
|[Application](documentitem-application-property-outlook.md)|Returns an  **[Application](application-object-outlook.md)** object that represents the parent Outlook application for the object. Read-only.|
|[Attachments](documentitem-attachments-property-outlook.md)|Returns an  **[Attachments](attachments-object-outlook.md)** object that represents all the attachments for the specified item. Read-only.|
|[AutoResolvedWinner](documentitem-autoresolvedwinner-property-outlook.md)|Returns a  **Boolean** that determines if the item is a winner of an automatic conflict resolution. Read-only.|
|[BillingInformation](documentitem-billinginformation-property-outlook.md)|Returns or sets a  **String** representing the billing information associated with the Outlook item. Read/write.|
|[Body](documentitem-body-property-outlook.md)|This object, member, or enumeration will be deprecated and is not intended to be used in your code. |
|[Categories](documentitem-categories-property-outlook.md)|Returns or sets a  **String** representing the categories assigned to the Outlook item. Read/write.|
|[Class](documentitem-class-property-outlook.md)|Returns an  **[OlObjectClass](olobjectclass-enumeration-outlook.md)** constant indicating the object's class. Read-only.|
|[Companies](documentitem-companies-property-outlook.md)|Returns or sets a  **String** representing the names of the companies associated with the Outlook item. Read/write.|
|[Conflicts](documentitem-conflicts-property-outlook.md)|Return the  **[Conflicts](conflicts-object-outlook.md)** object that represents the items that are in conflict for any Outlook item object. Read-only.|
|[ConversationIndex](documentitem-conversationindex-property-outlook.md)|Returns a  **String** that indicates the relative position of the item within the conversation thread. Read-only.|
|[ConversationTopic](documentitem-conversationtopic-property-outlook.md)|Returns a  **String** representing the topic of the conversation thread of the Outlook item. Read-only.|
|[CreationTime](documentitem-creationtime-property-outlook.md)|Returns a  **Date** indicating the creation time for the Outlook item. Read-only.|
|[DownloadState](documentitem-downloadstate-property-outlook.md)|Returns a constant that belongs to the  **[OlDownloadState](oldownloadstate-enumeration-outlook.md)** enumeration indicating the download state of the item. Read-only.|
|[EntryID](documentitem-entryid-property-outlook.md)|Returns a  **String** representing the unique Entry ID of the object. Read-only.|
|[FormDescription](documentitem-formdescription-property-outlook.md)|Returns the  **[FormDescription](formdescription-object-outlook.md)** object that represents the form description for the specified Outlook item. Read-only.|
|[GetInspector](documentitem-getinspector-property-outlook.md)|This object, member, or enumeration will be deprecated and is not intended to be used in your code. |
|[Importance](documentitem-importance-property-outlook.md)|Returns or sets an  **[OlImportance](olimportance-enumeration-outlook.md)** constant indicating the relative importance level for the Outlook item. Read/write.|
|[IsConflict](documentitem-isconflict-property-outlook.md)|Returns a  **Boolean** that determines if the item is in conflict. Read-only.|
|[ItemProperties](documentitem-itemproperties-property-outlook.md)|Returns an  **[ItemProperties](itemproperties-object-outlook.md)** collection that represents all standard and user-defined properties associated with the Outlook item. Read-only.|
|[LastModificationTime](documentitem-lastmodificationtime-property-outlook.md)|Returns a  **Date** specifying the date and time that the Outlook item was last modified. Read-only.|
|[MarkForDownload](documentitem-markfordownload-property-outlook.md)|This object, member, or enumeration will be deprecated and is not intended to be used in your code. |
|[MessageClass](documentitem-messageclass-property-outlook.md)|Returns or sets a  **String** representing the message class for the Outlook item. Read/write.|
|[Mileage](documentitem-mileage-property-outlook.md)|Returns or sets a  **String** representing the mileage for an item. Read/write.|
|[NoAging](documentitem-noaging-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** to not age the Outlook item. Read/write.|
|[OutlookInternalVersion](documentitem-outlookinternalversion-property-outlook.md)|Returns a  **Long** representing the build number of the Outlook application for an Outlook item. Read-only.|
|[OutlookVersion](documentitem-outlookversion-property-outlook.md)|Returns a  **String** indicating the major and minor version number of the Outlook application for an Outlook item. Read-only.|
|[Parent](documentitem-parent-property-outlook.md)|Returns the parent  **Object** of the specified object. Read-only.|
|[PropertyAccessor](documentitem-propertyaccessor-property-outlook.md)|Returns a  **[PropertyAccessor](propertyaccessor-object-outlook.md)** object that supports creating, getting, setting, and deleting properties of the parent **[DocumentItem](documentitem-object-outlook.md)** object. Read-only.|
|[Saved](documentitem-saved-property-outlook.md)|Returns a  **Boolean** value that is **True** if the Outlook item has not been modified since the last save. Read-only.|
|[Sensitivity](documentitem-sensitivity-property-outlook.md)|Returns or sets a constant in the  **[OlSensitivity](olsensitivity-enumeration-outlook.md)** enumeration indicating the sensitivity for the Outlook item. Read/write.|
|[Session](documentitem-session-property-outlook.md)|Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.|
|[Size](documentitem-size-property-outlook.md)|Returns a  **Long** indicating the size (in bytes) of the Outlook item. Read-only.|
|[Subject](documentitem-subject-property-outlook.md)|Returns or sets a  **String** indicating the subject for the Outlook item. Read/write.|
|[UnRead](documentitem-unread-property-outlook.md)|Returns or sets a  **Boolean** value that is **True** if the Outlook item has not been opened (read). Read/write.|
|[UserProperties](documentitem-userproperties-property-outlook.md)|Returns the  **[UserProperties](userproperties-object-outlook.md)** collection that represents all the user properties for the Outlook item. Read-only.|

