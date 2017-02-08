---
title: DocumentItem Events (Outlook)
ms.prod: OUTLOOK
ms.assetid: a0da5ec6-75ee-4fd6-bad3-22cf017df5b3
---


# DocumentItem Events (Outlook)
This object has the following events:

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

