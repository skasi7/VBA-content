---
title: MailItem Object (Outlook)
keywords: vbaol11.chm2987
f1_keywords:
- vbaol11.chm2987
ms.prod: OUTLOOK
api_name:
- Outlook.MailItem
ms.assetid: 14197346-05d2-0250-fa4c-4a6b07daf25f
---


# MailItem Object (Outlook)

Represents a mail message.


## Remarks

Use the  **[CreateItem](http://msdn.microsoft.com/library/application-createitem-method-outlook%28Office.15%29.aspx)** method to create a **MailItem** object that represents a new mail message.

Use the  **[Folder.Items](http://msdn.microsoft.com/library/folder-items-property-outlook%28Office.15%29.aspx)** property to obtain an **[Items](http://msdn.microsoft.com/library/items-object-outlook%28Office.15%29.aspx)** collection representing the mail items in a folder, and the **[Items.Item](http://msdn.microsoft.com/library/items-item-method-outlook%28Office.15%29.aspx)** (_index_) method, where _index_ is the index number of a mail message or a value used to match the default property of a message, to return a single **MailItem** object from the specified folder.


## Example

The following Visual Basic for Applications (VBA) example creates and displays a new mail message.


```
Sub CreateMail() 
 
 Dim myItem As Object 
 
 
 
 Set myItem = Application.CreateItem(olMailItem) 
 
 myItem.Subject = "Mail to myself" 
 
 myItem.Display 
 
End Sub
```

The following VBA example sets the current folder as the Inbox and displays the second mail message in the folder. In general, the order of mail messages in a folder is not guaranteed to be in a particular order. 




```
Sub DisplayMail() 
 
 Dim myItem As Object 
 
 Dim myFolder As Folder 
 
 
 
 Set myNamespace = Application.GetNamespace("MAPI") 
 
 Set myFolder = myNamespace.GetDefaultFolder(olFolderInbox) 
 
 myFolder.Display 
 
 Set myItem = myFolder.Items(2) 
 
 myItem.Display 
 
End Sub
```


## Events



|**Name**|
|:-----|
|[AfterWrite](http://msdn.microsoft.com/library/mailitem-afterwrite-event-outlook%28Office.15%29.aspx)|
|[AttachmentAdd](http://msdn.microsoft.com/library/mailitem-attachmentadd-event-outlook%28Office.15%29.aspx)|
|[AttachmentRead](http://msdn.microsoft.com/library/mailitem-attachmentread-event-outlook%28Office.15%29.aspx)|
|[AttachmentRemove](http://msdn.microsoft.com/library/mailitem-attachmentremove-event-outlook%28Office.15%29.aspx)|
|[BeforeAttachmentAdd](http://msdn.microsoft.com/library/mailitem-beforeattachmentadd-event-outlook%28Office.15%29.aspx)|
|[BeforeAttachmentPreview](http://msdn.microsoft.com/library/mailitem-beforeattachmentpreview-event-outlook%28Office.15%29.aspx)|
|[BeforeAttachmentRead](http://msdn.microsoft.com/library/mailitem-beforeattachmentread-event-outlook%28Office.15%29.aspx)|
|[BeforeAttachmentSave](http://msdn.microsoft.com/library/mailitem-beforeattachmentsave-event-outlook%28Office.15%29.aspx)|
|[BeforeAttachmentWriteToTempFile](http://msdn.microsoft.com/library/mailitem-beforeattachmentwritetotempfile-event-outlook%28Office.15%29.aspx)|
|[BeforeAutoSave](http://msdn.microsoft.com/library/mailitem-beforeautosave-event-outlook%28Office.15%29.aspx)|
|[BeforeCheckNames](http://msdn.microsoft.com/library/mailitem-beforechecknames-event-outlook%28Office.15%29.aspx)|
|[BeforeDelete](http://msdn.microsoft.com/library/mailitem-beforedelete-event-outlook%28Office.15%29.aspx)|
|[BeforeRead](http://msdn.microsoft.com/library/mailitem-beforeread-event-outlook%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/mailitem-close-event-outlook%28Office.15%29.aspx)|
|[CustomAction](http://msdn.microsoft.com/library/mailitem-customaction-event-outlook%28Office.15%29.aspx)|
|[CustomPropertyChange](http://msdn.microsoft.com/library/mailitem-custompropertychange-event-outlook%28Office.15%29.aspx)|
|[Forward](http://msdn.microsoft.com/library/mailitem-forward-event-outlook%28Office.15%29.aspx)|
|[Open](http://msdn.microsoft.com/library/mailitem-open-event-outlook%28Office.15%29.aspx)|
|[PropertyChange](http://msdn.microsoft.com/library/mailitem-propertychange-event-outlook%28Office.15%29.aspx)|
|[Read](http://msdn.microsoft.com/library/mailitem-read-event-outlook%28Office.15%29.aspx)|
|[ReadComplete](http://msdn.microsoft.com/library/mailitem-readcomplete-event-outlook%28Office.15%29.aspx)|
|[Reply](http://msdn.microsoft.com/library/mailitem-reply-event-outlook%28Office.15%29.aspx)|
|[ReplyAll](http://msdn.microsoft.com/library/mailitem-replyall-event-outlook%28Office.15%29.aspx)|
|[Send](http://msdn.microsoft.com/library/mailitem-send-event-outlook%28Office.15%29.aspx)|
|[Unload](http://msdn.microsoft.com/library/mailitem-unload-event-outlook%28Office.15%29.aspx)|
|[Write](http://msdn.microsoft.com/library/mailitem-write-event-outlook%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[AddBusinessCard](http://msdn.microsoft.com/library/mailitem-addbusinesscard-method-outlook%28Office.15%29.aspx)|
|[ClearConversationIndex](http://msdn.microsoft.com/library/mailitem-clearconversationindex-method-outlook%28Office.15%29.aspx)|
|[ClearTaskFlag](http://msdn.microsoft.com/library/mailitem-cleartaskflag-method-outlook%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/mailitem-close-method-outlook%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/mailitem-copy-method-outlook%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/mailitem-delete-method-outlook%28Office.15%29.aspx)|
|[Display](http://msdn.microsoft.com/library/mailitem-display-method-outlook%28Office.15%29.aspx)|
|[Forward](http://msdn.microsoft.com/library/mailitem-forward-method-outlook%28Office.15%29.aspx)|
|[GetConversation](http://msdn.microsoft.com/library/mailitem-getconversation-method-outlook%28Office.15%29.aspx)|
|[MarkAsTask](http://msdn.microsoft.com/library/mailitem-markastask-method-outlook%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/mailitem-move-method-outlook%28Office.15%29.aspx)|
|[PrintOut](http://msdn.microsoft.com/library/mailitem-printout-method-outlook%28Office.15%29.aspx)|
|[Reply](http://msdn.microsoft.com/library/mailitem-reply-method-outlook%28Office.15%29.aspx)|
|[ReplyAll](http://msdn.microsoft.com/library/mailitem-replyall-method-outlook%28Office.15%29.aspx)|
|[Save](http://msdn.microsoft.com/library/mailitem-save-method-outlook%28Office.15%29.aspx)|
|[SaveAs](http://msdn.microsoft.com/library/mailitem-saveas-method-outlook%28Office.15%29.aspx)|
|[Send](http://msdn.microsoft.com/library/mailitem-send-method-outlook%28Office.15%29.aspx)|
|[ShowCategoriesDialog](http://msdn.microsoft.com/library/mailitem-showcategoriesdialog-method-outlook%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Actions](http://msdn.microsoft.com/library/mailitem-actions-property-outlook%28Office.15%29.aspx)|
|[AlternateRecipientAllowed](http://msdn.microsoft.com/library/mailitem-alternaterecipientallowed-property-outlook%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/mailitem-application-property-outlook%28Office.15%29.aspx)|
|[Attachments](http://msdn.microsoft.com/library/mailitem-attachments-property-outlook%28Office.15%29.aspx)|
|[AutoForwarded](http://msdn.microsoft.com/library/mailitem-autoforwarded-property-outlook%28Office.15%29.aspx)|
|[AutoResolvedWinner](http://msdn.microsoft.com/library/mailitem-autoresolvedwinner-property-outlook%28Office.15%29.aspx)|
|[BCC](http://msdn.microsoft.com/library/mailitem-bcc-property-outlook%28Office.15%29.aspx)|
|[BillingInformation](http://msdn.microsoft.com/library/mailitem-billinginformation-property-outlook%28Office.15%29.aspx)|
|[Body](http://msdn.microsoft.com/library/mailitem-body-property-outlook%28Office.15%29.aspx)|
|[BodyFormat](http://msdn.microsoft.com/library/mailitem-bodyformat-property-outlook%28Office.15%29.aspx)|
|[Categories](http://msdn.microsoft.com/library/mailitem-categories-property-outlook%28Office.15%29.aspx)|
|[CC](http://msdn.microsoft.com/library/mailitem-cc-property-outlook%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/mailitem-class-property-outlook%28Office.15%29.aspx)|
|[Companies](http://msdn.microsoft.com/library/mailitem-companies-property-outlook%28Office.15%29.aspx)|
|[Conflicts](http://msdn.microsoft.com/library/mailitem-conflicts-property-outlook%28Office.15%29.aspx)|
|[ConversationID](http://msdn.microsoft.com/library/mailitem-conversationid-property-outlook%28Office.15%29.aspx)|
|[ConversationIndex](http://msdn.microsoft.com/library/mailitem-conversationindex-property-outlook%28Office.15%29.aspx)|
|[ConversationTopic](http://msdn.microsoft.com/library/mailitem-conversationtopic-property-outlook%28Office.15%29.aspx)|
|[CreationTime](http://msdn.microsoft.com/library/mailitem-creationtime-property-outlook%28Office.15%29.aspx)|
|[DeferredDeliveryTime](http://msdn.microsoft.com/library/mailitem-deferreddeliverytime-property-outlook%28Office.15%29.aspx)|
|[DeleteAfterSubmit](http://msdn.microsoft.com/library/mailitem-deleteaftersubmit-property-outlook%28Office.15%29.aspx)|
|[DownloadState](http://msdn.microsoft.com/library/mailitem-downloadstate-property-outlook%28Office.15%29.aspx)|
|[EntryID](http://msdn.microsoft.com/library/mailitem-entryid-property-outlook%28Office.15%29.aspx)|
|[ExpiryTime](http://msdn.microsoft.com/library/mailitem-expirytime-property-outlook%28Office.15%29.aspx)|
|[FlagRequest](http://msdn.microsoft.com/library/mailitem-flagrequest-property-outlook%28Office.15%29.aspx)|
|[FormDescription](http://msdn.microsoft.com/library/mailitem-formdescription-property-outlook%28Office.15%29.aspx)|
|[GetInspector](http://msdn.microsoft.com/library/mailitem-getinspector-property-outlook%28Office.15%29.aspx)|
|[HTMLBody](http://msdn.microsoft.com/library/mailitem-htmlbody-property-outlook%28Office.15%29.aspx)|
|[Importance](http://msdn.microsoft.com/library/mailitem-importance-property-outlook%28Office.15%29.aspx)|
|[InternetCodepage](http://msdn.microsoft.com/library/mailitem-internetcodepage-property-outlook%28Office.15%29.aspx)|
|[IsConflict](http://msdn.microsoft.com/library/mailitem-isconflict-property-outlook%28Office.15%29.aspx)|
|[IsMarkedAsTask](http://msdn.microsoft.com/library/mailitem-ismarkedastask-property-outlook%28Office.15%29.aspx)|
|[ItemProperties](http://msdn.microsoft.com/library/mailitem-itemproperties-property-outlook%28Office.15%29.aspx)|
|[LastModificationTime](http://msdn.microsoft.com/library/mailitem-lastmodificationtime-property-outlook%28Office.15%29.aspx)|
|[MarkForDownload](http://msdn.microsoft.com/library/mailitem-markfordownload-property-outlook%28Office.15%29.aspx)|
|[MessageClass](http://msdn.microsoft.com/library/mailitem-messageclass-property-outlook%28Office.15%29.aspx)|
|[Mileage](http://msdn.microsoft.com/library/mailitem-mileage-property-outlook%28Office.15%29.aspx)|
|[NoAging](http://msdn.microsoft.com/library/mailitem-noaging-property-outlook%28Office.15%29.aspx)|
|[OriginatorDeliveryReportRequested](http://msdn.microsoft.com/library/mailitem-originatordeliveryreportrequested-property-outlook%28Office.15%29.aspx)|
|[OutlookInternalVersion](http://msdn.microsoft.com/library/mailitem-outlookinternalversion-property-outlook%28Office.15%29.aspx)|
|[OutlookVersion](http://msdn.microsoft.com/library/mailitem-outlookversion-property-outlook%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/mailitem-parent-property-outlook%28Office.15%29.aspx)|
|[Permission](http://msdn.microsoft.com/library/mailitem-permission-property-outlook%28Office.15%29.aspx)|
|[PermissionService](http://msdn.microsoft.com/library/mailitem-permissionservice-property-outlook%28Office.15%29.aspx)|
|[PermissionTemplateGuid](http://msdn.microsoft.com/library/mailitem-permissiontemplateguid-property-outlook%28Office.15%29.aspx)|
|[PropertyAccessor](http://msdn.microsoft.com/library/mailitem-propertyaccessor-property-outlook%28Office.15%29.aspx)|
|[ReadReceiptRequested](http://msdn.microsoft.com/library/mailitem-readreceiptrequested-property-outlook%28Office.15%29.aspx)|
|[ReceivedByEntryID](http://msdn.microsoft.com/library/mailitem-receivedbyentryid-property-outlook%28Office.15%29.aspx)|
|[ReceivedByName](http://msdn.microsoft.com/library/mailitem-receivedbyname-property-outlook%28Office.15%29.aspx)|
|[ReceivedOnBehalfOfEntryID](http://msdn.microsoft.com/library/mailitem-receivedonbehalfofentryid-property-outlook%28Office.15%29.aspx)|
|[ReceivedOnBehalfOfName](http://msdn.microsoft.com/library/mailitem-receivedonbehalfofname-property-outlook%28Office.15%29.aspx)|
|[ReceivedTime](http://msdn.microsoft.com/library/mailitem-receivedtime-property-outlook%28Office.15%29.aspx)|
|[RecipientReassignmentProhibited](http://msdn.microsoft.com/library/mailitem-recipientreassignmentprohibited-property-outlook%28Office.15%29.aspx)|
|[Recipients](http://msdn.microsoft.com/library/mailitem-recipients-property-outlook%28Office.15%29.aspx)|
|[ReminderOverrideDefault](http://msdn.microsoft.com/library/mailitem-reminderoverridedefault-property-outlook%28Office.15%29.aspx)|
|[ReminderPlaySound](http://msdn.microsoft.com/library/mailitem-reminderplaysound-property-outlook%28Office.15%29.aspx)|
|[ReminderSet](http://msdn.microsoft.com/library/mailitem-reminderset-property-outlook%28Office.15%29.aspx)|
|[ReminderSoundFile](http://msdn.microsoft.com/library/mailitem-remindersoundfile-property-outlook%28Office.15%29.aspx)|
|[ReminderTime](http://msdn.microsoft.com/library/mailitem-remindertime-property-outlook%28Office.15%29.aspx)|
|[RemoteStatus](http://msdn.microsoft.com/library/mailitem-remotestatus-property-outlook%28Office.15%29.aspx)|
|[ReplyRecipientNames](http://msdn.microsoft.com/library/mailitem-replyrecipientnames-property-outlook%28Office.15%29.aspx)|
|[ReplyRecipients](http://msdn.microsoft.com/library/mailitem-replyrecipients-property-outlook%28Office.15%29.aspx)|
|[RetentionExpirationDate](http://msdn.microsoft.com/library/mailitem-retentionexpirationdate-property-outlook%28Office.15%29.aspx)|
|[RetentionPolicyName](http://msdn.microsoft.com/library/mailitem-retentionpolicyname-property-outlook%28Office.15%29.aspx)|
|[RTFBody](http://msdn.microsoft.com/library/mailitem-rtfbody-property-outlook%28Office.15%29.aspx)|
|[Saved](http://msdn.microsoft.com/library/mailitem-saved-property-outlook%28Office.15%29.aspx)|
|[SaveSentMessageFolder](http://msdn.microsoft.com/library/mailitem-savesentmessagefolder-property-outlook%28Office.15%29.aspx)|
|[Sender](http://msdn.microsoft.com/library/mailitem-sender-property-outlook%28Office.15%29.aspx)|
|[SenderEmailAddress](http://msdn.microsoft.com/library/mailitem-senderemailaddress-property-outlook%28Office.15%29.aspx)|
|[SenderEmailType](http://msdn.microsoft.com/library/mailitem-senderemailtype-property-outlook%28Office.15%29.aspx)|
|[SenderName](http://msdn.microsoft.com/library/mailitem-sendername-property-outlook%28Office.15%29.aspx)|
|[SendUsingAccount](http://msdn.microsoft.com/library/mailitem-sendusingaccount-property-outlook%28Office.15%29.aspx)|
|[Sensitivity](http://msdn.microsoft.com/library/mailitem-sensitivity-property-outlook%28Office.15%29.aspx)|
|[Sent](http://msdn.microsoft.com/library/mailitem-sent-property-outlook%28Office.15%29.aspx)|
|[SentOn](http://msdn.microsoft.com/library/mailitem-senton-property-outlook%28Office.15%29.aspx)|
|[SentOnBehalfOfName](http://msdn.microsoft.com/library/mailitem-sentonbehalfofname-property-outlook%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/mailitem-session-property-outlook%28Office.15%29.aspx)|
|[Size](http://msdn.microsoft.com/library/mailitem-size-property-outlook%28Office.15%29.aspx)|
|[Subject](http://msdn.microsoft.com/library/mailitem-subject-property-outlook%28Office.15%29.aspx)|
|[Submitted](http://msdn.microsoft.com/library/mailitem-submitted-property-outlook%28Office.15%29.aspx)|
|[TaskCompletedDate](http://msdn.microsoft.com/library/mailitem-taskcompleteddate-property-outlook%28Office.15%29.aspx)|
|[TaskDueDate](http://msdn.microsoft.com/library/mailitem-taskduedate-property-outlook%28Office.15%29.aspx)|
|[TaskStartDate](http://msdn.microsoft.com/library/mailitem-taskstartdate-property-outlook%28Office.15%29.aspx)|
|[TaskSubject](http://msdn.microsoft.com/library/mailitem-tasksubject-property-outlook%28Office.15%29.aspx)|
|[To](http://msdn.microsoft.com/library/mailitem-to-property-outlook%28Office.15%29.aspx)|
|[ToDoTaskOrdinal](http://msdn.microsoft.com/library/mailitem-todotaskordinal-property-outlook%28Office.15%29.aspx)|
|[UnRead](http://msdn.microsoft.com/library/mailitem-unread-property-outlook%28Office.15%29.aspx)|
|[UserProperties](http://msdn.microsoft.com/library/mailitem-userproperties-property-outlook%28Office.15%29.aspx)|
|[VotingOptions](http://msdn.microsoft.com/library/mailitem-votingoptions-property-outlook%28Office.15%29.aspx)|
|[VotingResponse](http://msdn.microsoft.com/library/mailitem-votingresponse-property-outlook%28Office.15%29.aspx)|

## See also


#### Other resources


[Send an E-mail Given the SMTP Address of an Account (Outlook)](http://msdn.microsoft.com/library/send-an-e-mail-given-the-smtp-address-of-an-account-outlook%28Office.15%29.aspx)<br>
[Outlook Object Model Reference](http://msdn.microsoft.com/library/object-model-outlook-vba-reference%28Office.15%29.aspx)
<<<<<<< HEAD
=======

>>>>>>> d7667e83d23dbf8ebf5bf068ba6fed14c840c0f5

