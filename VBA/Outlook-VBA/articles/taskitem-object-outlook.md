---
title: TaskItem Object (Outlook)
keywords: vbaol11.chm2990
f1_keywords:
- vbaol11.chm2990
ms.prod: OUTLOOK
api_name:
- Outlook.TaskItem
ms.assetid: 5df8cfa5-5460-a5a1-a130-ba5bca1a0091
---


# TaskItem Object (Outlook)

Represents a task (an assigned, delegated, or self-imposed task to be performed within a specified time frame) in a Tasks folder.


## Remarks

Use the  **[CreateItem](http://msdn.microsoft.com/library/application-createitem-method-outlook%28Office.15%29.aspx)** method to create a **TaskItem** object that represents a new task.

Use  **[Items](http://msdn.microsoft.com/library/folder-items-property-outlook%28Office.15%29.aspx)** ( _index_ ), where _index_ is the index number of a task or a value used to match the default property of a task, to return a single **TaskItem** object from a Tasks folder.


## Example

The following Visual Basic for Applications (VBA) example returns a new task.






```
Set myItem = Application.CreateItem(olTaskItem)
```


## Events



|**Name**|
|:-----|
|[AfterWrite](http://msdn.microsoft.com/library/taskitem-afterwrite-event-outlook%28Office.15%29.aspx)|
|[AttachmentAdd](http://msdn.microsoft.com/library/taskitem-attachmentadd-event-outlook%28Office.15%29.aspx)|
|[AttachmentRead](http://msdn.microsoft.com/library/taskitem-attachmentread-event-outlook%28Office.15%29.aspx)|
|[AttachmentRemove](http://msdn.microsoft.com/library/taskitem-attachmentremove-event-outlook%28Office.15%29.aspx)|
|[BeforeAttachmentAdd](http://msdn.microsoft.com/library/taskitem-beforeattachmentadd-event-outlook%28Office.15%29.aspx)|
|[BeforeAttachmentPreview](http://msdn.microsoft.com/library/taskitem-beforeattachmentpreview-event-outlook%28Office.15%29.aspx)|
|[BeforeAttachmentRead](http://msdn.microsoft.com/library/taskitem-beforeattachmentread-event-outlook%28Office.15%29.aspx)|
|[BeforeAttachmentSave](http://msdn.microsoft.com/library/taskitem-beforeattachmentsave-event-outlook%28Office.15%29.aspx)|
|[BeforeAttachmentWriteToTempFile](http://msdn.microsoft.com/library/taskitem-beforeattachmentwritetotempfile-event-outlook%28Office.15%29.aspx)|
|[BeforeAutoSave](http://msdn.microsoft.com/library/taskitem-beforeautosave-event-outlook%28Office.15%29.aspx)|
|[BeforeCheckNames](http://msdn.microsoft.com/library/taskitem-beforechecknames-event-outlook%28Office.15%29.aspx)|
|[BeforeDelete](http://msdn.microsoft.com/library/taskitem-beforedelete-event-outlook%28Office.15%29.aspx)|
|[BeforeRead](http://msdn.microsoft.com/library/taskitem-beforeread-event-outlook%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/taskitem-close-event-outlook%28Office.15%29.aspx)|
|[CustomAction](http://msdn.microsoft.com/library/taskitem-customaction-event-outlook%28Office.15%29.aspx)|
|[CustomPropertyChange](http://msdn.microsoft.com/library/taskitem-custompropertychange-event-outlook%28Office.15%29.aspx)|
|[Forward](http://msdn.microsoft.com/library/taskitem-forward-event-outlook%28Office.15%29.aspx)|
|[Open](http://msdn.microsoft.com/library/taskitem-open-event-outlook%28Office.15%29.aspx)|
|[PropertyChange](http://msdn.microsoft.com/library/taskitem-propertychange-event-outlook%28Office.15%29.aspx)|
|[Read](http://msdn.microsoft.com/library/taskitem-read-event-outlook%28Office.15%29.aspx)|
|[ReadComplete](http://msdn.microsoft.com/library/taskitem-readcomplete-event-outlook%28Office.15%29.aspx)|
|[Reply](http://msdn.microsoft.com/library/taskitem-reply-event-outlook%28Office.15%29.aspx)|
|[ReplyAll](http://msdn.microsoft.com/library/taskitem-replyall-event-outlook%28Office.15%29.aspx)|
|[Send](http://msdn.microsoft.com/library/taskitem-send-event-outlook%28Office.15%29.aspx)|
|[Unload](http://msdn.microsoft.com/library/taskitem-unload-event-outlook%28Office.15%29.aspx)|
|[Write](http://msdn.microsoft.com/library/taskitem-write-event-outlook%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Assign](http://msdn.microsoft.com/library/taskitem-assign-method-outlook%28Office.15%29.aspx)|
|[CancelResponseState](http://msdn.microsoft.com/library/taskitem-cancelresponsestate-method-outlook%28Office.15%29.aspx)|
|[ClearRecurrencePattern](http://msdn.microsoft.com/library/taskitem-clearrecurrencepattern-method-outlook%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/taskitem-close-method-outlook%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/taskitem-copy-method-outlook%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/taskitem-delete-method-outlook%28Office.15%29.aspx)|
|[Display](http://msdn.microsoft.com/library/taskitem-display-method-outlook%28Office.15%29.aspx)|
|[GetConversation](http://msdn.microsoft.com/library/taskitem-getconversation-method-outlook%28Office.15%29.aspx)|
|[GetRecurrencePattern](http://msdn.microsoft.com/library/taskitem-getrecurrencepattern-method-outlook%28Office.15%29.aspx)|
|[MarkComplete](http://msdn.microsoft.com/library/taskitem-markcomplete-method-outlook%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/taskitem-move-method-outlook%28Office.15%29.aspx)|
|[PrintOut](http://msdn.microsoft.com/library/taskitem-printout-method-outlook%28Office.15%29.aspx)|
|[Respond](http://msdn.microsoft.com/library/taskitem-respond-method-outlook%28Office.15%29.aspx)|
|[Save](http://msdn.microsoft.com/library/taskitem-save-method-outlook%28Office.15%29.aspx)|
|[SaveAs](http://msdn.microsoft.com/library/taskitem-saveas-method-outlook%28Office.15%29.aspx)|
|[Send](http://msdn.microsoft.com/library/taskitem-send-method-outlook%28Office.15%29.aspx)|
|[ShowCategoriesDialog](http://msdn.microsoft.com/library/taskitem-showcategoriesdialog-method-outlook%28Office.15%29.aspx)|
|[SkipRecurrence](http://msdn.microsoft.com/library/taskitem-skiprecurrence-method-outlook%28Office.15%29.aspx)|
|[StatusReport](http://msdn.microsoft.com/library/taskitem-statusreport-method-outlook%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Actions](http://msdn.microsoft.com/library/taskitem-actions-property-outlook%28Office.15%29.aspx)|
|[ActualWork](http://msdn.microsoft.com/library/taskitem-actualwork-property-outlook%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/taskitem-application-property-outlook%28Office.15%29.aspx)|
|[Attachments](http://msdn.microsoft.com/library/taskitem-attachments-property-outlook%28Office.15%29.aspx)|
|[AutoResolvedWinner](http://msdn.microsoft.com/library/taskitem-autoresolvedwinner-property-outlook%28Office.15%29.aspx)|
|[BillingInformation](http://msdn.microsoft.com/library/taskitem-billinginformation-property-outlook%28Office.15%29.aspx)|
|[Body](http://msdn.microsoft.com/library/taskitem-body-property-outlook%28Office.15%29.aspx)|
|[CardData](http://msdn.microsoft.com/library/taskitem-carddata-property-outlook%28Office.15%29.aspx)|
|[Categories](http://msdn.microsoft.com/library/taskitem-categories-property-outlook%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/taskitem-class-property-outlook%28Office.15%29.aspx)|
|[Companies](http://msdn.microsoft.com/library/taskitem-companies-property-outlook%28Office.15%29.aspx)|
|[Complete](http://msdn.microsoft.com/library/taskitem-complete-property-outlook%28Office.15%29.aspx)|
|[Conflicts](http://msdn.microsoft.com/library/taskitem-conflicts-property-outlook%28Office.15%29.aspx)|
|[ContactNames](http://msdn.microsoft.com/library/taskitem-contactnames-property-outlook%28Office.15%29.aspx)|
|[ConversationID](http://msdn.microsoft.com/library/taskitem-conversationid-property-outlook%28Office.15%29.aspx)|
|[ConversationIndex](http://msdn.microsoft.com/library/taskitem-conversationindex-property-outlook%28Office.15%29.aspx)|
|[ConversationTopic](http://msdn.microsoft.com/library/taskitem-conversationtopic-property-outlook%28Office.15%29.aspx)|
|[CreationTime](http://msdn.microsoft.com/library/taskitem-creationtime-property-outlook%28Office.15%29.aspx)|
|[DateCompleted](http://msdn.microsoft.com/library/taskitem-datecompleted-property-outlook%28Office.15%29.aspx)|
|[DelegationState](http://msdn.microsoft.com/library/taskitem-delegationstate-property-outlook%28Office.15%29.aspx)|
|[Delegator](http://msdn.microsoft.com/library/taskitem-delegator-property-outlook%28Office.15%29.aspx)|
|[DownloadState](http://msdn.microsoft.com/library/taskitem-downloadstate-property-outlook%28Office.15%29.aspx)|
|[DueDate](http://msdn.microsoft.com/library/taskitem-duedate-property-outlook%28Office.15%29.aspx)|
|[EntryID](http://msdn.microsoft.com/library/taskitem-entryid-property-outlook%28Office.15%29.aspx)|
|[FormDescription](http://msdn.microsoft.com/library/taskitem-formdescription-property-outlook%28Office.15%29.aspx)|
|[GetInspector](http://msdn.microsoft.com/library/taskitem-getinspector-property-outlook%28Office.15%29.aspx)|
|[Importance](http://msdn.microsoft.com/library/taskitem-importance-property-outlook%28Office.15%29.aspx)|
|[InternetCodepage](http://msdn.microsoft.com/library/taskitem-internetcodepage-property-outlook%28Office.15%29.aspx)|
|[IsConflict](http://msdn.microsoft.com/library/taskitem-isconflict-property-outlook%28Office.15%29.aspx)|
|[IsRecurring](http://msdn.microsoft.com/library/taskitem-isrecurring-property-outlook%28Office.15%29.aspx)|
|[ItemProperties](http://msdn.microsoft.com/library/taskitem-itemproperties-property-outlook%28Office.15%29.aspx)|
|[LastModificationTime](http://msdn.microsoft.com/library/taskitem-lastmodificationtime-property-outlook%28Office.15%29.aspx)|
|[MarkForDownload](http://msdn.microsoft.com/library/taskitem-markfordownload-property-outlook%28Office.15%29.aspx)|
|[MessageClass](http://msdn.microsoft.com/library/taskitem-messageclass-property-outlook%28Office.15%29.aspx)|
|[Mileage](http://msdn.microsoft.com/library/taskitem-mileage-property-outlook%28Office.15%29.aspx)|
|[NoAging](http://msdn.microsoft.com/library/taskitem-noaging-property-outlook%28Office.15%29.aspx)|
|[Ordinal](http://msdn.microsoft.com/library/taskitem-ordinal-property-outlook%28Office.15%29.aspx)|
|[OutlookInternalVersion](http://msdn.microsoft.com/library/taskitem-outlookinternalversion-property-outlook%28Office.15%29.aspx)|
|[OutlookVersion](http://msdn.microsoft.com/library/taskitem-outlookversion-property-outlook%28Office.15%29.aspx)|
|[Owner](http://msdn.microsoft.com/library/taskitem-owner-property-outlook%28Office.15%29.aspx)|
|[Ownership](http://msdn.microsoft.com/library/taskitem-ownership-property-outlook%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/taskitem-parent-property-outlook%28Office.15%29.aspx)|
|[PercentComplete](http://msdn.microsoft.com/library/taskitem-percentcomplete-property-outlook%28Office.15%29.aspx)|
|[PropertyAccessor](http://msdn.microsoft.com/library/taskitem-propertyaccessor-property-outlook%28Office.15%29.aspx)|
|[Recipients](http://msdn.microsoft.com/library/taskitem-recipients-property-outlook%28Office.15%29.aspx)|
|[ReminderOverrideDefault](http://msdn.microsoft.com/library/taskitem-reminderoverridedefault-property-outlook%28Office.15%29.aspx)|
|[ReminderPlaySound](http://msdn.microsoft.com/library/taskitem-reminderplaysound-property-outlook%28Office.15%29.aspx)|
|[ReminderSet](http://msdn.microsoft.com/library/taskitem-reminderset-property-outlook%28Office.15%29.aspx)|
|[ReminderSoundFile](http://msdn.microsoft.com/library/taskitem-remindersoundfile-property-outlook%28Office.15%29.aspx)|
|[ReminderTime](http://msdn.microsoft.com/library/taskitem-remindertime-property-outlook%28Office.15%29.aspx)|
|[ResponseState](http://msdn.microsoft.com/library/taskitem-responsestate-property-outlook%28Office.15%29.aspx)|
|[Role](http://msdn.microsoft.com/library/taskitem-role-property-outlook%28Office.15%29.aspx)|
|[RTFBody](http://msdn.microsoft.com/library/taskitem-rtfbody-property-outlook%28Office.15%29.aspx)|
|[Saved](http://msdn.microsoft.com/library/taskitem-saved-property-outlook%28Office.15%29.aspx)|
|[SchedulePlusPriority](http://msdn.microsoft.com/library/taskitem-schedulepluspriority-property-outlook%28Office.15%29.aspx)|
|[SendUsingAccount](http://msdn.microsoft.com/library/taskitem-sendusingaccount-property-outlook%28Office.15%29.aspx)|
|[Sensitivity](http://msdn.microsoft.com/library/taskitem-sensitivity-property-outlook%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/taskitem-session-property-outlook%28Office.15%29.aspx)|
|[Size](http://msdn.microsoft.com/library/taskitem-size-property-outlook%28Office.15%29.aspx)|
|[StartDate](http://msdn.microsoft.com/library/taskitem-startdate-property-outlook%28Office.15%29.aspx)|
|[Status](http://msdn.microsoft.com/library/taskitem-status-property-outlook%28Office.15%29.aspx)|
|[StatusOnCompletionRecipients](http://msdn.microsoft.com/library/taskitem-statusoncompletionrecipients-property-outlook%28Office.15%29.aspx)|
|[StatusUpdateRecipients](http://msdn.microsoft.com/library/taskitem-statusupdaterecipients-property-outlook%28Office.15%29.aspx)|
|[Subject](http://msdn.microsoft.com/library/taskitem-subject-property-outlook%28Office.15%29.aspx)|
|[TeamTask](http://msdn.microsoft.com/library/taskitem-teamtask-property-outlook%28Office.15%29.aspx)|
|[ToDoTaskOrdinal](http://msdn.microsoft.com/library/taskitem-todotaskordinal-property-outlook%28Office.15%29.aspx)|
|[TotalWork](http://msdn.microsoft.com/library/taskitem-totalwork-property-outlook%28Office.15%29.aspx)|
|[UnRead](http://msdn.microsoft.com/library/taskitem-unread-property-outlook%28Office.15%29.aspx)|
|[UserProperties](http://msdn.microsoft.com/library/taskitem-userproperties-property-outlook%28Office.15%29.aspx)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/object-model-outlook-vba-reference%28Office.15%29.aspx)
[TaskItem Object Members](http://msdn.microsoft.com/library/taskitem-members-outlook%28Office.15%29.aspx)
