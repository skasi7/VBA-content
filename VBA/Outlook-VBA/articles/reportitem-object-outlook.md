---
title: ReportItem Object (Outlook)
keywords: vbaol11.chm3007
f1_keywords:
- vbaol11.chm3007
ms.prod: OUTLOOK
api_name:
- Outlook.ReportItem
ms.assetid: 16ebe336-72e0-42f6-99d3-edecc3ea284d
---


# ReportItem Object (Outlook)

Represents a mail-delivery report in an Inbox folder. 


## Remarks

The  **ReportItem** object is similar to a **[MailItem](http://msdn.microsoft.com/library/mailitem-object-outlook%28Office.15%29.aspx)** object, and it contains a report (usually the non-delivery report) or error message from the mail transport system.

Unlike other Microsoft Outlook objects, you cannot create this object. Report items are created automatically when any report or error in general is received from the mail transport system.


## Events



|**Name**|
|:-----|
|[AfterWrite](http://msdn.microsoft.com/library/reportitem-afterwrite-event-outlook%28Office.15%29.aspx)|
|[AttachmentAdd](http://msdn.microsoft.com/library/reportitem-attachmentadd-event-outlook%28Office.15%29.aspx)|
|[AttachmentRead](http://msdn.microsoft.com/library/reportitem-attachmentread-event-outlook%28Office.15%29.aspx)|
|[AttachmentRemove](http://msdn.microsoft.com/library/reportitem-attachmentremove-event-outlook%28Office.15%29.aspx)|
|[BeforeAttachmentAdd](http://msdn.microsoft.com/library/reportitem-beforeattachmentadd-event-outlook%28Office.15%29.aspx)|
|[BeforeAttachmentPreview](http://msdn.microsoft.com/library/reportitem-beforeattachmentpreview-event-outlook%28Office.15%29.aspx)|
|[BeforeAttachmentRead](http://msdn.microsoft.com/library/reportitem-beforeattachmentread-event-outlook%28Office.15%29.aspx)|
|[BeforeAttachmentSave](http://msdn.microsoft.com/library/reportitem-beforeattachmentsave-event-outlook%28Office.15%29.aspx)|
|[BeforeAttachmentWriteToTempFile](http://msdn.microsoft.com/library/reportitem-beforeattachmentwritetotempfile-event-outlook%28Office.15%29.aspx)|
|[BeforeAutoSave](http://msdn.microsoft.com/library/reportitem-beforeautosave-event-outlook%28Office.15%29.aspx)|
|[BeforeCheckNames](http://msdn.microsoft.com/library/reportitem-beforechecknames-event-outlook%28Office.15%29.aspx)|
|[BeforeDelete](http://msdn.microsoft.com/library/reportitem-beforedelete-event-outlook%28Office.15%29.aspx)|
|[BeforeRead](http://msdn.microsoft.com/library/reportitem-beforeread-event-outlook%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/reportitem-close-event-outlook%28Office.15%29.aspx)|
|[CustomAction](http://msdn.microsoft.com/library/reportitem-customaction-event-outlook%28Office.15%29.aspx)|
|[CustomPropertyChange](http://msdn.microsoft.com/library/reportitem-custompropertychange-event-outlook%28Office.15%29.aspx)|
|[Forward](http://msdn.microsoft.com/library/reportitem-forward-event-outlook%28Office.15%29.aspx)|
|[Open](http://msdn.microsoft.com/library/reportitem-open-event-outlook%28Office.15%29.aspx)|
|[PropertyChange](http://msdn.microsoft.com/library/reportitem-propertychange-event-outlook%28Office.15%29.aspx)|
|[Read](http://msdn.microsoft.com/library/reportitem-read-event-outlook%28Office.15%29.aspx)|
|[ReadComplete](http://msdn.microsoft.com/library/reportitem-readcomplete-event-outlook%28Office.15%29.aspx)|
|[Reply](http://msdn.microsoft.com/library/reportitem-reply-event-outlook%28Office.15%29.aspx)|
|[ReplyAll](http://msdn.microsoft.com/library/reportitem-replyall-event-outlook%28Office.15%29.aspx)|
|[Send](http://msdn.microsoft.com/library/reportitem-send-event-outlook%28Office.15%29.aspx)|
|[Unload](http://msdn.microsoft.com/library/reportitem-unload-event-outlook%28Office.15%29.aspx)|
|[Write](http://msdn.microsoft.com/library/reportitem-write-event-outlook%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Close](http://msdn.microsoft.com/library/reportitem-close-method-outlook%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/reportitem-copy-method-outlook%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/reportitem-delete-method-outlook%28Office.15%29.aspx)|
|[Display](http://msdn.microsoft.com/library/reportitem-display-method-outlook%28Office.15%29.aspx)|
|[GetConversation](http://msdn.microsoft.com/library/reportitem-getconversation-method-outlook%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/reportitem-move-method-outlook%28Office.15%29.aspx)|
|[PrintOut](http://msdn.microsoft.com/library/reportitem-printout-method-outlook%28Office.15%29.aspx)|
|[Save](http://msdn.microsoft.com/library/reportitem-save-method-outlook%28Office.15%29.aspx)|
|[SaveAs](http://msdn.microsoft.com/library/reportitem-saveas-method-outlook%28Office.15%29.aspx)|
|[ShowCategoriesDialog](http://msdn.microsoft.com/library/reportitem-showcategoriesdialog-method-outlook%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Actions](http://msdn.microsoft.com/library/reportitem-actions-property-outlook%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/reportitem-application-property-outlook%28Office.15%29.aspx)|
|[Attachments](http://msdn.microsoft.com/library/reportitem-attachments-property-outlook%28Office.15%29.aspx)|
|[AutoResolvedWinner](http://msdn.microsoft.com/library/reportitem-autoresolvedwinner-property-outlook%28Office.15%29.aspx)|
|[BillingInformation](http://msdn.microsoft.com/library/reportitem-billinginformation-property-outlook%28Office.15%29.aspx)|
|[Body](http://msdn.microsoft.com/library/reportitem-body-property-outlook%28Office.15%29.aspx)|
|[Categories](http://msdn.microsoft.com/library/reportitem-categories-property-outlook%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/reportitem-class-property-outlook%28Office.15%29.aspx)|
|[Companies](http://msdn.microsoft.com/library/reportitem-companies-property-outlook%28Office.15%29.aspx)|
|[Conflicts](http://msdn.microsoft.com/library/reportitem-conflicts-property-outlook%28Office.15%29.aspx)|
|[ConversationID](http://msdn.microsoft.com/library/reportitem-conversationid-property-outlook%28Office.15%29.aspx)|
|[ConversationIndex](http://msdn.microsoft.com/library/reportitem-conversationindex-property-outlook%28Office.15%29.aspx)|
|[ConversationTopic](http://msdn.microsoft.com/library/reportitem-conversationtopic-property-outlook%28Office.15%29.aspx)|
|[CreationTime](http://msdn.microsoft.com/library/reportitem-creationtime-property-outlook%28Office.15%29.aspx)|
|[DownloadState](http://msdn.microsoft.com/library/reportitem-downloadstate-property-outlook%28Office.15%29.aspx)|
|[EntryID](http://msdn.microsoft.com/library/reportitem-entryid-property-outlook%28Office.15%29.aspx)|
|[FormDescription](http://msdn.microsoft.com/library/reportitem-formdescription-property-outlook%28Office.15%29.aspx)|
|[GetInspector](http://msdn.microsoft.com/library/reportitem-getinspector-property-outlook%28Office.15%29.aspx)|
|[Importance](http://msdn.microsoft.com/library/reportitem-importance-property-outlook%28Office.15%29.aspx)|
|[IsConflict](http://msdn.microsoft.com/library/reportitem-isconflict-property-outlook%28Office.15%29.aspx)|
|[ItemProperties](http://msdn.microsoft.com/library/reportitem-itemproperties-property-outlook%28Office.15%29.aspx)|
|[LastModificationTime](http://msdn.microsoft.com/library/reportitem-lastmodificationtime-property-outlook%28Office.15%29.aspx)|
|[MarkForDownload](http://msdn.microsoft.com/library/reportitem-markfordownload-property-outlook%28Office.15%29.aspx)|
|[MessageClass](http://msdn.microsoft.com/library/reportitem-messageclass-property-outlook%28Office.15%29.aspx)|
|[Mileage](http://msdn.microsoft.com/library/reportitem-mileage-property-outlook%28Office.15%29.aspx)|
|[NoAging](http://msdn.microsoft.com/library/reportitem-noaging-property-outlook%28Office.15%29.aspx)|
|[OutlookInternalVersion](http://msdn.microsoft.com/library/reportitem-outlookinternalversion-property-outlook%28Office.15%29.aspx)|
|[OutlookVersion](http://msdn.microsoft.com/library/reportitem-outlookversion-property-outlook%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/reportitem-parent-property-outlook%28Office.15%29.aspx)|
|[PropertyAccessor](http://msdn.microsoft.com/library/reportitem-propertyaccessor-property-outlook%28Office.15%29.aspx)|
|[RetentionExpirationDate](http://msdn.microsoft.com/library/reportitem-retentionexpirationdate-property-outlook%28Office.15%29.aspx)|
|[RetentionPolicyName](http://msdn.microsoft.com/library/reportitem-retentionpolicyname-property-outlook%28Office.15%29.aspx)|
|[Saved](http://msdn.microsoft.com/library/reportitem-saved-property-outlook%28Office.15%29.aspx)|
|[Sensitivity](http://msdn.microsoft.com/library/reportitem-sensitivity-property-outlook%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/reportitem-session-property-outlook%28Office.15%29.aspx)|
|[Size](http://msdn.microsoft.com/library/reportitem-size-property-outlook%28Office.15%29.aspx)|
|[Subject](http://msdn.microsoft.com/library/reportitem-subject-property-outlook%28Office.15%29.aspx)|
|[UnRead](http://msdn.microsoft.com/library/reportitem-unread-property-outlook%28Office.15%29.aspx)|
|[UserProperties](http://msdn.microsoft.com/library/reportitem-userproperties-property-outlook%28Office.15%29.aspx)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/object-model-outlook-vba-reference%28Office.15%29.aspx)
[ReportItem Object Members](http://msdn.microsoft.com/library/reportitem-members-outlook%28Office.15%29.aspx)
