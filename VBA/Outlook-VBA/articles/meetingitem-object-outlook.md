---
title: MeetingItem Object (Outlook)
keywords: vbaol11.chm2989
f1_keywords:
- vbaol11.chm2989
ms.prod: OUTLOOK
api_name:
- Outlook.MeetingItem
ms.assetid: b75730f5-b395-3d66-5acd-b64fd8fcd78f
---


# MeetingItem Object (Outlook)

Represents a change to the recipient's Calendar folder initiated by another party or as a result of a group action.


## Remarks

Unlike other Microsoft Outlook objects, you cannot create this object. It is created automatically when you set the  **[MeetingStatus](http://msdn.microsoft.com/library/appointmentitem-meetingstatus-property-outlook%28Office.15%29.aspx)** property of an **[AppointmentItem](appointmentitem-object-outlook.md)** object to **olMeeting** and send it to one or more users. They receive it in their inboxes as a **MeetingItem**.

Use the  **[GetAssociatedAppointment](http://msdn.microsoft.com/library/meetingitem-getassociatedappointment-method-outlook%28Office.15%29.aspx)** method to return the **AppointmentItem** object associated with a **MeetingItem** object, and work directly with the **AppointmentItem** object to respond to the request.


## Example

The following example uses the  **[CreateItem](http://msdn.microsoft.com/library/application-createitem-method-outlook%28Office.15%29.aspx)** method to create an appointment. It becomes a **MeetingItem** with both a required and an optional attendee when it is received in the inbox of each of the recipients.


```
Set myItem = myOlApp.CreateItem(olAppointmentItem) 
 
myItem.MeetingStatus = olMeeting 
 
myItem.Subject = "Strategy Meeting" 
 
myItem.Location = "Conference Room B" 
 
myItem.Start = #9/24/97 1:30:00 PM# 
 
myItem.Duration = 90 
 
Set myRequiredAttendee = myItem.Recipients.Add("Nate _ 
 
 Sun") 
 
myRequiredAttendee.Type = olRequired 
 
Set myOptionalAttendee = myItem.Recipients.Add("Kevin _ 
 
 Kennedy") 
 
myOptionalAttendee.Type = olOptional 
 
Set myResourceAttendee = _ 
 
 myItem.Recipients.Add("Conference Room B") 
 
myResourceAttendee.Type = olResource 
 
myItem.Send
```


## Events



|**Name**|
|:-----|
|[AfterWrite](http://msdn.microsoft.com/library/meetingitem-afterwrite-event-outlook%28Office.15%29.aspx)|
|[AttachmentAdd](http://msdn.microsoft.com/library/meetingitem-attachmentadd-event-outlook%28Office.15%29.aspx)|
|[AttachmentRead](http://msdn.microsoft.com/library/meetingitem-attachmentread-event-outlook%28Office.15%29.aspx)|
|[AttachmentRemove](http://msdn.microsoft.com/library/meetingitem-attachmentremove-event-outlook%28Office.15%29.aspx)|
|[BeforeAttachmentAdd](http://msdn.microsoft.com/library/meetingitem-beforeattachmentadd-event-outlook%28Office.15%29.aspx)|
|[BeforeAttachmentPreview](http://msdn.microsoft.com/library/meetingitem-beforeattachmentpreview-event-outlook%28Office.15%29.aspx)|
|[BeforeAttachmentRead](http://msdn.microsoft.com/library/meetingitem-beforeattachmentread-event-outlook%28Office.15%29.aspx)|
|[BeforeAttachmentSave](http://msdn.microsoft.com/library/meetingitem-beforeattachmentsave-event-outlook%28Office.15%29.aspx)|
|[BeforeAttachmentWriteToTempFile](http://msdn.microsoft.com/library/meetingitem-beforeattachmentwritetotempfile-event-outlook%28Office.15%29.aspx)|
|[BeforeAutoSave](http://msdn.microsoft.com/library/meetingitem-beforeautosave-event-outlook%28Office.15%29.aspx)|
|[BeforeCheckNames](http://msdn.microsoft.com/library/meetingitem-beforechecknames-event-outlook%28Office.15%29.aspx)|
|[BeforeDelete](http://msdn.microsoft.com/library/meetingitem-beforedelete-event-outlook%28Office.15%29.aspx)|
|[BeforeRead](http://msdn.microsoft.com/library/meetingitem-beforeread-event-outlook%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/meetingitem-close-event-outlook%28Office.15%29.aspx)|
|[CustomAction](http://msdn.microsoft.com/library/meetingitem-customaction-event-outlook%28Office.15%29.aspx)|
|[CustomPropertyChange](http://msdn.microsoft.com/library/meetingitem-custompropertychange-event-outlook%28Office.15%29.aspx)|
|[Forward](http://msdn.microsoft.com/library/meetingitem-forward-event-outlook%28Office.15%29.aspx)|
|[Open](http://msdn.microsoft.com/library/meetingitem-open-event-outlook%28Office.15%29.aspx)|
|[PropertyChange](http://msdn.microsoft.com/library/meetingitem-propertychange-event-outlook%28Office.15%29.aspx)|
|[Read](http://msdn.microsoft.com/library/meetingitem-read-event-outlook%28Office.15%29.aspx)|
|[ReadComplete](http://msdn.microsoft.com/library/meetingitem-readcomplete-event-outlook%28Office.15%29.aspx)|
|[Reply](http://msdn.microsoft.com/library/meetingitem-reply-event-outlook%28Office.15%29.aspx)|
|[ReplyAll](http://msdn.microsoft.com/library/meetingitem-replyall-event-outlook%28Office.15%29.aspx)|
|[Send](http://msdn.microsoft.com/library/meetingitem-send-event-outlook%28Office.15%29.aspx)|
|[Unload](http://msdn.microsoft.com/library/meetingitem-unload-event-outlook%28Office.15%29.aspx)|
|[Write](http://msdn.microsoft.com/library/meetingitem-write-event-outlook%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Close](http://msdn.microsoft.com/library/meetingitem-close-method-outlook%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/meetingitem-copy-method-outlook%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/meetingitem-delete-method-outlook%28Office.15%29.aspx)|
|[Display](http://msdn.microsoft.com/library/meetingitem-display-method-outlook%28Office.15%29.aspx)|
|[Forward](http://msdn.microsoft.com/library/meetingitem-forward-method-outlook%28Office.15%29.aspx)|
|[GetAssociatedAppointment](http://msdn.microsoft.com/library/meetingitem-getassociatedappointment-method-outlook%28Office.15%29.aspx)|
|[GetConversation](http://msdn.microsoft.com/library/meetingitem-getconversation-method-outlook%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/meetingitem-move-method-outlook%28Office.15%29.aspx)|
|[PrintOut](http://msdn.microsoft.com/library/meetingitem-printout-method-outlook%28Office.15%29.aspx)|
|[Reply](http://msdn.microsoft.com/library/meetingitem-reply-method-outlook%28Office.15%29.aspx)|
|[ReplyAll](http://msdn.microsoft.com/library/meetingitem-replyall-method-outlook%28Office.15%29.aspx)|
|[Save](http://msdn.microsoft.com/library/meetingitem-save-method-outlook%28Office.15%29.aspx)|
|[SaveAs](http://msdn.microsoft.com/library/meetingitem-saveas-method-outlook%28Office.15%29.aspx)|
|[Send](http://msdn.microsoft.com/library/meetingitem-send-method-outlook%28Office.15%29.aspx)|
|[ShowCategoriesDialog](http://msdn.microsoft.com/library/meetingitem-showcategoriesdialog-method-outlook%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Actions](http://msdn.microsoft.com/library/meetingitem-actions-property-outlook%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/meetingitem-application-property-outlook%28Office.15%29.aspx)|
|[Attachments](http://msdn.microsoft.com/library/meetingitem-attachments-property-outlook%28Office.15%29.aspx)|
|[AutoForwarded](http://msdn.microsoft.com/library/meetingitem-autoforwarded-property-outlook%28Office.15%29.aspx)|
|[AutoResolvedWinner](http://msdn.microsoft.com/library/meetingitem-autoresolvedwinner-property-outlook%28Office.15%29.aspx)|
|[BillingInformation](http://msdn.microsoft.com/library/meetingitem-billinginformation-property-outlook%28Office.15%29.aspx)|
|[Body](http://msdn.microsoft.com/library/meetingitem-body-property-outlook%28Office.15%29.aspx)|
|[Categories](http://msdn.microsoft.com/library/meetingitem-categories-property-outlook%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/meetingitem-class-property-outlook%28Office.15%29.aspx)|
|[Companies](http://msdn.microsoft.com/library/meetingitem-companies-property-outlook%28Office.15%29.aspx)|
|[Conflicts](http://msdn.microsoft.com/library/meetingitem-conflicts-property-outlook%28Office.15%29.aspx)|
|[ConversationID](http://msdn.microsoft.com/library/meetingitem-conversationid-property-outlook%28Office.15%29.aspx)|
|[ConversationIndex](http://msdn.microsoft.com/library/meetingitem-conversationindex-property-outlook%28Office.15%29.aspx)|
|[ConversationTopic](http://msdn.microsoft.com/library/meetingitem-conversationtopic-property-outlook%28Office.15%29.aspx)|
|[CreationTime](http://msdn.microsoft.com/library/meetingitem-creationtime-property-outlook%28Office.15%29.aspx)|
|[DeferredDeliveryTime](http://msdn.microsoft.com/library/meetingitem-deferreddeliverytime-property-outlook%28Office.15%29.aspx)|
|[DeleteAfterSubmit](http://msdn.microsoft.com/library/meetingitem-deleteaftersubmit-property-outlook%28Office.15%29.aspx)|
|[DownloadState](http://msdn.microsoft.com/library/meetingitem-downloadstate-property-outlook%28Office.15%29.aspx)|
|[EntryID](http://msdn.microsoft.com/library/meetingitem-entryid-property-outlook%28Office.15%29.aspx)|
|[ExpiryTime](http://msdn.microsoft.com/library/meetingitem-expirytime-property-outlook%28Office.15%29.aspx)|
|[FormDescription](http://msdn.microsoft.com/library/meetingitem-formdescription-property-outlook%28Office.15%29.aspx)|
|[GetInspector](http://msdn.microsoft.com/library/meetingitem-getinspector-property-outlook%28Office.15%29.aspx)|
|[Importance](http://msdn.microsoft.com/library/meetingitem-importance-property-outlook%28Office.15%29.aspx)|
|[IsConflict](http://msdn.microsoft.com/library/meetingitem-isconflict-property-outlook%28Office.15%29.aspx)|
|[IsLatestVersion](http://msdn.microsoft.com/library/meetingitem-islatestversion-property-outlook%28Office.15%29.aspx)|
|[ItemProperties](http://msdn.microsoft.com/library/meetingitem-itemproperties-property-outlook%28Office.15%29.aspx)|
|[LastModificationTime](http://msdn.microsoft.com/library/meetingitem-lastmodificationtime-property-outlook%28Office.15%29.aspx)|
|[MarkForDownload](http://msdn.microsoft.com/library/meetingitem-markfordownload-property-outlook%28Office.15%29.aspx)|
|[MeetingWorkspaceURL](http://msdn.microsoft.com/library/meetingitem-meetingworkspaceurl-property-outlook%28Office.15%29.aspx)|
|[MessageClass](http://msdn.microsoft.com/library/meetingitem-messageclass-property-outlook%28Office.15%29.aspx)|
|[Mileage](http://msdn.microsoft.com/library/meetingitem-mileage-property-outlook%28Office.15%29.aspx)|
|[NoAging](http://msdn.microsoft.com/library/meetingitem-noaging-property-outlook%28Office.15%29.aspx)|
|[OriginatorDeliveryReportRequested](http://msdn.microsoft.com/library/meetingitem-originatordeliveryreportrequested-property-outlook%28Office.15%29.aspx)|
|[OutlookInternalVersion](http://msdn.microsoft.com/library/meetingitem-outlookinternalversion-property-outlook%28Office.15%29.aspx)|
|[OutlookVersion](http://msdn.microsoft.com/library/meetingitem-outlookversion-property-outlook%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/meetingitem-parent-property-outlook%28Office.15%29.aspx)|
|[PropertyAccessor](http://msdn.microsoft.com/library/meetingitem-propertyaccessor-property-outlook%28Office.15%29.aspx)|
|[ReceivedTime](http://msdn.microsoft.com/library/meetingitem-receivedtime-property-outlook%28Office.15%29.aspx)|
|[Recipients](http://msdn.microsoft.com/library/meetingitem-recipients-property-outlook%28Office.15%29.aspx)|
|[ReminderSet](http://msdn.microsoft.com/library/meetingitem-reminderset-property-outlook%28Office.15%29.aspx)|
|[ReminderTime](http://msdn.microsoft.com/library/meetingitem-remindertime-property-outlook%28Office.15%29.aspx)|
|[ReplyRecipients](http://msdn.microsoft.com/library/meetingitem-replyrecipients-property-outlook%28Office.15%29.aspx)|
|[RetentionExpirationDate](http://msdn.microsoft.com/library/meetingitem-retentionexpirationdate-property-outlook%28Office.15%29.aspx)|
|[RetentionPolicyName](http://msdn.microsoft.com/library/meetingitem-retentionpolicyname-property-outlook%28Office.15%29.aspx)|
|[RTFBody](http://msdn.microsoft.com/library/meetingitem-rtfbody-property-outlook%28Office.15%29.aspx)|
|[Saved](http://msdn.microsoft.com/library/meetingitem-saved-property-outlook%28Office.15%29.aspx)|
|[SaveSentMessageFolder](http://msdn.microsoft.com/library/meetingitem-savesentmessagefolder-property-outlook%28Office.15%29.aspx)|
|[SenderEmailAddress](http://msdn.microsoft.com/library/meetingitem-senderemailaddress-property-outlook%28Office.15%29.aspx)|
|[SenderEmailType](http://msdn.microsoft.com/library/meetingitem-senderemailtype-property-outlook%28Office.15%29.aspx)|
|[SenderName](http://msdn.microsoft.com/library/meetingitem-sendername-property-outlook%28Office.15%29.aspx)|
|[SendUsingAccount](http://msdn.microsoft.com/library/meetingitem-sendusingaccount-property-outlook%28Office.15%29.aspx)|
|[Sensitivity](http://msdn.microsoft.com/library/meetingitem-sensitivity-property-outlook%28Office.15%29.aspx)|
|[Sent](http://msdn.microsoft.com/library/meetingitem-sent-property-outlook%28Office.15%29.aspx)|
|[SentOn](http://msdn.microsoft.com/library/meetingitem-senton-property-outlook%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/meetingitem-session-property-outlook%28Office.15%29.aspx)|
|[Size](http://msdn.microsoft.com/library/meetingitem-size-property-outlook%28Office.15%29.aspx)|
|[Subject](http://msdn.microsoft.com/library/meetingitem-subject-property-outlook%28Office.15%29.aspx)|
|[Submitted](http://msdn.microsoft.com/library/meetingitem-submitted-property-outlook%28Office.15%29.aspx)|
|[UnRead](http://msdn.microsoft.com/library/meetingitem-unread-property-outlook%28Office.15%29.aspx)|
|[UserProperties](http://msdn.microsoft.com/library/meetingitem-userproperties-property-outlook%28Office.15%29.aspx)|

## See also


#### Other resources


[MeetingItem Object Members](http://msdn.microsoft.com/library/meetingitem-members-outlook%28Office.15%29.aspx)
[Outlook Object Model Reference](http://msdn.microsoft.com/library/object-model-outlook-vba-reference%28Office.15%29.aspx)
